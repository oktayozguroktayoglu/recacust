document.documentElement.classList.add('js');
const reduceMotion = window.matchMedia('(prefers-reduced-motion: reduce)').matches;

// Masaüstü mega menü: fare ile üzerine gelince ya da klavyeyle odaklanınca açılır, Esc ile kapanır.
const header = document.getElementById('mainHeader');
const nav = document.getElementById('desktopNav');
if (header && nav) {
  let hideTimer;
  const panels = [...document.querySelectorAll('.mega-panel[data-panel]')];
  const triggers = [...nav.querySelectorAll('[data-mega]')];
  const closeMega = () => {
    header.classList.remove('mega-open');
    panels.forEach(p => { p.classList.remove('active'); p.setAttribute('aria-hidden', 'true'); });
    triggers.forEach(t => { if (t.dataset.mega !== '0') t.setAttribute('aria-expanded', 'false'); });
  };
  const openMega = (name) => {
    clearTimeout(hideTimer);
    panels.forEach(p => {
      const active = p.dataset.panel === name;
      p.classList.toggle('active', active);
      p.setAttribute('aria-hidden', active ? 'false' : 'true');
    });
    triggers.forEach(t => { if (t.dataset.mega !== '0') t.setAttribute('aria-expanded', t.dataset.mega === name ? 'true' : 'false'); });
    header.classList.add('mega-open');
  };
  triggers.forEach(item => {
    const name = item.dataset.mega;
    if (name && name !== '0') {
      item.setAttribute('aria-expanded', 'false');
      item.addEventListener('mouseenter', () => openMega(name));
      item.addEventListener('focus', () => openMega(name));
    } else {
      item.addEventListener('mouseenter', closeMega);
      item.addEventListener('focus', closeMega);
    }
  });
  panels.forEach(panel => panel.addEventListener('mouseenter', () => clearTimeout(hideTimer)));
  header.addEventListener('mouseleave', () => { hideTimer = setTimeout(closeMega, 140); });
  header.addEventListener('focusout', (e) => { if (!header.contains(e.relatedTarget)) closeMega(); });
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeMega(); });
  const onScroll = () => header.classList.toggle('scrolled', window.scrollY > 8);
  onScroll();
  window.addEventListener('scroll', onScroll, { passive: true });
}

// Mobil menü
const openBtn = document.getElementById('mobileOpen');
const closeBtn = document.getElementById('mobileClose');
const drawer = document.getElementById('mobileDrawer');
const backdrop = document.getElementById('drawerBackdrop');
if (drawer && openBtn) {
  const setDrawer = (open) => {
    drawer.classList.toggle('open', open);
    backdrop?.classList.toggle('open', open);
    drawer.setAttribute('aria-hidden', open ? 'false' : 'true');
    drawer.inert = !open;
    openBtn.setAttribute('aria-expanded', open ? 'true' : 'false');
    document.body.style.overflow = open ? 'hidden' : '';
    (open ? closeBtn : openBtn)?.focus({ preventScroll: true });
  };
  drawer.inert = true;
  openBtn.setAttribute('aria-expanded', 'false');
  openBtn.addEventListener('click', () => setDrawer(true));
  closeBtn?.addEventListener('click', () => setDrawer(false));
  backdrop?.addEventListener('click', () => setDrawer(false));
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && drawer.classList.contains('open')) setDrawer(false); });
  // Aynı sayfadaki bir başlığa giden bağlantıya basılınca menü kapanır.
  drawer.querySelectorAll('a[href*="#"]').forEach(a => a.addEventListener('click', () => {
    drawer.classList.remove('open'); backdrop?.classList.remove('open');
    drawer.inert = true; drawer.setAttribute('aria-hidden', 'true');
    openBtn.setAttribute('aria-expanded', 'false'); document.body.style.overflow = '';
  }));
  const groups = [...drawer.querySelectorAll('.drawer-body details')];
  groups.forEach(d => d.addEventListener('toggle', () => { if (d.open) groups.forEach(o => { if (o !== d) o.open = false; }); }));
}

// İçindekiler: okunan bölümü vurgular; mobilde çubuğu o bölüme kaydırır.
const toc = document.querySelector('.doc-toc');
if (toc && 'IntersectionObserver' in window) {
  const links = [...toc.querySelectorAll('a[href^="#"]')];
  const byId = new Map(links.map(a => [a.getAttribute('href').slice(1), a]));
  const setActive = (id) => {
    const link = byId.get(id);
    if (!link || link.classList.contains('is-active')) return;
    links.forEach(a => { a.classList.remove('is-active'); a.removeAttribute('aria-current'); });
    link.classList.add('is-active');
    link.setAttribute('aria-current', 'true');
    if (toc.scrollWidth > toc.clientWidth) {
      toc.scrollTo({ left: link.offsetLeft - (toc.clientWidth - link.offsetWidth) / 2, behavior: reduceMotion ? 'auto' : 'smooth' });
    }
  };
  const spy = new IntersectionObserver((entries) => {
    entries.forEach(en => { if (en.isIntersecting) setActive(en.target.id); });
  }, { rootMargin: '-30% 0px -60% 0px' });
  byId.forEach((_, id) => { const el = document.getElementById(id); if (el) spy.observe(el); });
}

// Yukarı çık düğmesi
const toTop = document.getElementById('toTop');
if (toTop) {
  const toggle = () => toTop.classList.toggle('show', window.scrollY > 900);
  toggle();
  window.addEventListener('scroll', toggle, { passive: true });
  toTop.addEventListener('click', () => window.scrollTo({ top: 0, behavior: reduceMotion ? 'auto' : 'smooth' }));
}

// Kartların görünür olunca hafifçe belirmesi
if ('IntersectionObserver' in window && !reduceMotion) {
  const items = document.querySelectorAll('.three-cards > div, .service-grid > div, .two-col-v5 > div, .home-list > li, .explore-card, .sector-grid-v5 > div, .hero-card-v5, .consult-grid');
  const io = new IntersectionObserver((entries) => {
    entries.forEach(en => { if (en.isIntersecting) { en.target.classList.add('in'); io.unobserve(en.target); } });
  }, { rootMargin: '0px 0px -8% 0px' });
  items.forEach((el, i) => { el.classList.add('reveal'); el.style.transitionDelay = (i % 4) * 60 + 'ms'; io.observe(el); });
}

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
      contactForm.classList.add('was-validated');
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
