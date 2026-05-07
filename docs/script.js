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

// Order form (Formspree, with Kakao/phone fallback when not configured)
document.getElementById('orderForm')?.addEventListener('submit', async e => {
  const form = e.target;
  e.preventDefault();
  if (!form.action || form.action.includes('YOUR_FORM_ID')) {
    const naverUrl = form.dataset.naverUrl;
    if (naverUrl) {
      window.open(naverUrl, '_blank', 'noopener');
      alert('주문은 네이버 폼에서 받고 있어요. 새 창에서 폼을 열었으니 그곳에 다시 작성해 주세요 🙏');
    } else {
      alert('주문 폼은 준비 중입니다. 빠른 문의는 카카오톡 채널(@EditO) 또는 010-6238-1934 로 부탁드려요 🍪');
    }
    return;
  }
  const btn = form.querySelector('button[type="submit"]');
  const original = btn.textContent;
  btn.disabled = true;
  btn.textContent = '보내는 중…';
  try {
    const res = await fetch(form.action, {
      method: 'POST',
      body: new FormData(form),
      headers: { 'Accept': 'application/json' }
    });
    if (res.ok) {
      alert('문의가 접수되었습니다. 영업일 1일 이내에 연락드릴게요 🍪');
      form.reset();
    } else {
      const data = await res.json().catch(() => ({}));
      alert(data?.errors?.[0]?.message || '전송에 실패했습니다. 잠시 후 다시 시도해 주세요.');
    }
  } catch {
    alert('네트워크 오류로 전송하지 못했습니다. 카카오톡(@EditO)으로 문의 부탁드립니다.');
  } finally {
    btn.disabled = false;
    btn.textContent = original;
  }
});
