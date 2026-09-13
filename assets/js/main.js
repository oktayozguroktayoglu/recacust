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
