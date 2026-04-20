// Mobile nav toggle
const navToggle = document.getElementById('navToggle');
const navMobile = document.getElementById('navMobile');
if (navToggle && navMobile) {
  navToggle.addEventListener('click', () => navMobile.classList.toggle('open'));
  navMobile.querySelectorAll('a').forEach(a =>
    a.addEventListener('click', () => navMobile.classList.remove('open'))
  );
}

// Subtle parallax on the hero cookie stack
const stack = document.querySelector('.cookie-stack');
if (stack && window.matchMedia('(pointer:fine)').matches) {
  stack.addEventListener('mousemove', (e) => {
    const rect = stack.getBoundingClientRect();
    const cx = (e.clientX - rect.left) / rect.width - 0.5;
    const cy = (e.clientY - rect.top) / rect.height - 0.5;
    stack.querySelectorAll('.cookie').forEach((el, i) => {
      const depth = (i % 3 + 1) * 6;
      el.style.transform =
        `translate(calc(var(--x) + ${cx * depth}px), calc(var(--y) + ${cy * depth}px)) rotate(var(--r))`;
    });
  });
  stack.addEventListener('mouseleave', () => {
    stack.querySelectorAll('.cookie').forEach((el) => {
      el.style.transform = '';
    });
  });
}

// Reveal sections on scroll
const io = new IntersectionObserver((entries) => {
  entries.forEach(e => {
    if (e.isIntersecting) {
      e.target.style.opacity = 1;
      e.target.style.transform = 'translateY(0)';
      io.unobserve(e.target);
    }
  });
}, { threshold: 0.1 });

document.querySelectorAll('.product, .ing, blockquote, .story__text, .story__image, .gift__visual, .gift__text').forEach(el => {
  el.style.opacity = 0;
  el.style.transform = 'translateY(24px)';
  el.style.transition = 'opacity .7s ease, transform .7s ease';
  io.observe(el);
});

// Order form — demo only
const form = document.getElementById('orderForm');
if (form) {
  form.addEventListener('submit', (e) => {
    e.preventDefault();
    const data = Object.fromEntries(new FormData(form).entries());
    alert(`문의 접수 완료 (데모):\n\n이름: ${data.name}\n연락처: ${data.phone}\n내용: ${data.message}\n\n실제 연동 전입니다. 010-6238-1934 로도 문의 주세요.`);
    form.reset();
  });
}
