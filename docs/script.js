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

// Order form demo
document.getElementById('orderForm')?.addEventListener('submit', e => {
  e.preventDefault();
  const d = Object.fromEntries(new FormData(e.target));
  alert(`문의 접수!\n\n이름: ${d.name}\n연락처: ${d.phone}\n\n빠른 시일 내 연락드리겠습니다 🍪\n(데모 — 실제 문의: 010-6238-1934)`);
  e.target.reset();
});
