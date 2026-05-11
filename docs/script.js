// Mobile nav
const burger = document.getElementById('burger');
const drawer = document.getElementById('drawer');
if (burger && drawer) {
  burger.addEventListener('click', () => drawer.classList.toggle('open'));
  drawer.querySelectorAll('a').forEach(a => a.addEventListener('click', () => drawer.classList.remove('open')));
}

// Scroll reveal
const ro = new IntersectionObserver(entries => {
  entries.forEach(e => { if (e.isIntersecting) { e.target.classList.add('visible'); ro.unobserve(e.target); } });
}, { threshold: 0.12 });

document.querySelectorAll('.pcard,.ing__item,blockquote,.story__text,.story__img,.gift__text,.gift__img,.gal__item,.story__nums > div').forEach(el => {
  el.classList.add('reveal');
  ro.observe(el);
});

// Nav shadow on scroll
window.addEventListener('scroll', () => {
  document.getElementById('nav')?.classList.toggle('nav--scrolled', window.scrollY > 40);
});

// Order tabs
const tabWrap = document.getElementById('orderTabs');
const tabToggle = document.getElementById('tabToggle');
if (tabWrap && tabToggle) {
  tabToggle.addEventListener('click', () => tabWrap.classList.toggle('open'));
  tabWrap.querySelectorAll('.order__tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      tabWrap.querySelectorAll('.order__tab-btn').forEach(b => b.classList.remove('active'));
      tabWrap.querySelectorAll('.order__tab-panel').forEach(p => p.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById(btn.dataset.tab)?.classList.add('active');
      if (!tabWrap.classList.contains('open')) tabWrap.classList.add('open');
    });
  });
}
