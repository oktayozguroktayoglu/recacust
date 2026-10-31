const header=document.getElementById('mainHeader');
const nav=document.getElementById('desktopNav');
if(header&&nav){
  let hideTimer;
  const panels=[...document.querySelectorAll('.mega-panel[data-panel]')];
  const closeMega=()=>{header.classList.remove('mega-open');panels.forEach(p=>{p.classList.remove('active');p.setAttribute('aria-hidden','true')});};
  const openMega=(name)=>{
    clearTimeout(hideTimer);
    panels.forEach(p=>{
      const active=p.dataset.panel===name;
      p.classList.toggle('active',active);
      p.setAttribute('aria-hidden',active?'false':'true');
    });
    header.classList.add('mega-open');
  };
  nav.querySelectorAll('[data-mega]').forEach(item=>{
    const name=item.dataset.mega;
    if(name&&name!=='0'){
      item.addEventListener('mouseenter',()=>openMega(name));
      item.addEventListener('focus',()=>openMega(name));
    }else{
      item.addEventListener('mouseenter',closeMega);
      item.addEventListener('focus',closeMega);
    }
  });
  panels.forEach(panel=>{
    panel.addEventListener('mouseenter',()=>clearTimeout(hideTimer));
  });
  header.addEventListener('mouseleave',()=>{hideTimer=setTimeout(closeMega,120)});
}
const openBtn=document.getElementById('mobileOpen');
const closeBtn=document.getElementById('mobileClose');
const drawer=document.getElementById('mobileDrawer');
const backdrop=document.getElementById('drawerBackdrop');
function openDrawer(){drawer?.classList.add('open');backdrop?.classList.add('open');document.body.style.overflow='hidden'}
function closeDrawer(){drawer?.classList.remove('open');backdrop?.classList.remove('open');document.body.style.overflow=''}
openBtn?.addEventListener('click',openDrawer);closeBtn?.addEventListener('click',closeDrawer);backdrop?.addEventListener('click',closeDrawer);
document.addEventListener('keydown',e=>{if(e.key==='Escape')closeDrawer()});
document.querySelectorAll('.drawer-body details').forEach(d=>{d.addEventListener('toggle',()=>{if(d.open){document.querySelectorAll('.drawer-body details').forEach(o=>{if(o!==d)o.open=false})}})});

// İletişim formu
// GitHub Pages PHP çalıştıramadığı için form Web3Forms (https://web3forms.com) üzerinden gönderilir.
// WEB3FORMS_KEY boşken form, bilgileri hazır bir e-posta olarak açar (yedek yöntem).
const WEB3FORMS_KEY = '';
const CONTACT_EMAIL = 'info@recacust.com';
const contactForm = document.getElementById('contactForm');
if (contactForm) {
  const status = document.getElementById('formStatus');
  const button = contactForm.querySelector('button[type="submit"]');
  const setStatus = (msg, type) => { if (status) { status.textContent = msg; status.dataset.type = type || ''; } };
  contactForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    if (!contactForm.checkValidity()) {
      contactForm.reportValidity();
      return;
    }
    const data = Object.fromEntries(new FormData(contactForm).entries());
    if (data.website) return; // bot tuzağı doldurulmuşsa sessizce yok say
    const lines = [
      'Ad Soyad: ' + data.name,
      'Kuruluş: ' + (data.company || '-'),
      'E-posta: ' + data.email,
      'Telefon: ' + (data.phone || '-'),
      'Talep konusu: ' + data.topic,
      '',
      data.message,
    ];
    if (!WEB3FORMS_KEY) {
      const subject = 'Web sitesi talebi: ' + data.topic + ' - ' + data.name;
      window.location.href = 'mailto:' + CONTACT_EMAIL + '?subject=' + encodeURIComponent(subject) + '&body=' + encodeURIComponent(lines.join('\n'));
      setStatus('E-posta uygulamanız açıldı. Mesajı göndererek talebinizi iletebilirsiniz. Açılmadıysa ' + CONTACT_EMAIL + ' adresine yazabilirsiniz.', 'info');
      return;
    }
    button.disabled = true;
    const label = button.textContent;
    button.textContent = 'Gönderiliyor…';
    setStatus('', '');
    try {
      const res = await fetch('https://api.web3forms.com/submit', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
        body: JSON.stringify({
          access_key: WEB3FORMS_KEY,
          subject: 'RECACUST web sitesi: ' + data.topic,
          from_name: 'RECACUST Web Sitesi',
          replyto: data.email,
          'Ad Soyad': data.name, 'Kuruluş': data.company || '-', 'E-posta': data.email,
          'Telefon': data.phone || '-', 'Talep konusu': data.topic, 'Mesaj': data.message, 'KVKK onayı': 'Evet',
        }),
      });
      const json = await res.json();
      if (!json.success) throw new Error(json.message || 'Gönderilemedi');
      window.location.href = 'tesekkurler.html';
    } catch (err) {
      setStatus('Mesajınız gönderilemedi. Lütfen tekrar deneyin ya da ' + CONTACT_EMAIL + ' adresine e-posta gönderin.', 'error');
      button.disabled = false;
      button.textContent = label;
    }
  });
}
