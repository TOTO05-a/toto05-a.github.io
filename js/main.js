(() => {
  'use strict';

  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => [...r.querySelectorAll(s)];
  const T = window.TRANSLATIONS || {};
  const reduceMotion = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  const DEFAULT_LANG = 'es';
  const DEFAULT_THEME = 'dark';

  /* ==================== IDIOMA ==================== */
  let lang = DEFAULT_LANG;
  try {
    const saved = localStorage.getItem('lang');
    if (saved && T[saved]) lang = saved;
  } catch (_) { /* localStorage no disponible */ }

  let theme = DEFAULT_THEME;
  try {
    const savedTheme = localStorage.getItem('theme');
    if (savedTheme === 'light' || savedTheme === 'dark') theme = savedTheme;
  } catch (_) { /* localStorage no disponible */ }

  const t = (key) => (T[lang] && T[lang][key] !== undefined ? T[lang][key] : key);

  const themeButton = $('#theme-toggle');
  function updateThemeButton() {
    if (!themeButton) return;
    const label = t(theme === 'dark' ? 'a11y.themeLight' : 'a11y.themeDark');
    themeButton.setAttribute('aria-label', label);
    themeButton.setAttribute('title', label);
  }

  function applyTheme(next) {
    if (next !== 'light' && next !== 'dark') return;
    theme = next;
    document.documentElement.dataset.theme = theme;
    try { localStorage.setItem('theme', theme); } catch (_) {}
    updateThemeButton();
  }

  if (themeButton) themeButton.addEventListener('click', () => applyTheme(theme === 'dark' ? 'light' : 'dark'));

  function applyLanguage(next) {
    if (!T[next]) return;
    lang = next;
    try { localStorage.setItem('lang', lang); } catch (_) {}

    document.documentElement.lang = lang;
    document.title = t('meta.title');
    const desc = $('meta[name="description"]');
    if (desc) desc.setAttribute('content', t('meta.desc'));
    $$('[data-i18n-content]').forEach((el) => { el.setAttribute('content', t(el.dataset.i18nContent)); });

    $$('[data-i18n]').forEach((el) => { el.textContent = t(el.dataset.i18n); });
    $$('[data-i18n-aria]').forEach((el) => { el.setAttribute('aria-label', t(el.dataset.i18nAria)); });
    $$('[data-i18n-title]').forEach((el) => { el.setAttribute('title', t(el.dataset.i18nTitle)); });
    $$('.lang button').forEach((b) => b.setAttribute('aria-pressed', String(b.dataset.lang === lang)));
    updateBurgerLabel();
    updateThemeButton();
    startTyping();
  }

  $$('.lang button').forEach((b) => b.addEventListener('click', () => applyLanguage(b.dataset.lang)));

  /* ==================== EFECTO DE ESCRITURA ==================== */
  const typingEl = $('#typing');
  let typingTimer;
  function startTyping() {
    clearTimeout(typingTimer);
    const words = t('hero.traits');
    if (!typingEl || !Array.isArray(words)) return;
    if (reduceMotion) { typingEl.textContent = words[0]; return; }

    let w = 0, c = 0, deleting = false;
    typingEl.textContent = '';
    (function tick() {
      const word = words[w];
      if (!deleting) {
        c++;
        typingEl.textContent = word.slice(0, c);
        if (c === word.length) { deleting = true; typingTimer = setTimeout(tick, 1800); return; }
        typingTimer = setTimeout(tick, 90);
      } else {
        c--;
        typingEl.textContent = word.slice(0, c);
        if (c === 0) { deleting = false; w = (w + 1) % words.length; typingTimer = setTimeout(tick, 300); return; }
        typingTimer = setTimeout(tick, 45);
      }
    })();
  }

  /* ==================== NAVBAR ==================== */
  const nav = $('#nav');
  const menu = $('#menu');
  const burger = $('#burger');

  function updateBurgerLabel() {
    const open = menu.classList.contains('is-open');
    burger.setAttribute('aria-label', t(open ? 'a11y.menuClose' : 'a11y.menu'));
  }
  function setMenu(open) {
    menu.classList.toggle('is-open', open);
    nav.classList.toggle('is-open', open);
    document.body.classList.toggle('menu-open', open);
    burger.setAttribute('aria-expanded', String(open));
    updateBurgerLabel();
  }
  burger.addEventListener('click', () => setMenu(!menu.classList.contains('is-open')));
  $$('a', menu).forEach((a, i) => {
    a.style.setProperty('--i', i);
    a.addEventListener('click', () => setMenu(false));
  });
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape') setMenu(false); });
  window.addEventListener('resize', () => { if (window.innerWidth > 860) setMenu(false); });

  /* Fondo del navbar + manchas del fondo (una sola llamada por frame) */
  const blobs = $$('.bg__blob');
  let ticking = false;
  function onScroll() {
    const y = window.scrollY;
    nav.classList.toggle('is-scrolled', y > 20);
    if (!reduceMotion) {
      blobs.forEach((b, i) => {
        b.style.setProperty('--bx', Math.sin(y / 500 + i * 1.7) * 120 + 'px');
        b.style.setProperty('--by', Math.cos(y / 600 + i) * 50 + 'px');
      });
    }
    ticking = false;
  }
  window.addEventListener('scroll', () => {
    if (!ticking) { ticking = true; requestAnimationFrame(onScroll); }
  }, { passive: true });
  onScroll();

  /* Enlace activo según la sección visible */
  if ('IntersectionObserver' in window) {
    const links = $$('.nav__menu a');
    const spy = new IntersectionObserver((entries) => {
      entries.forEach((e) => {
        if (e.isIntersecting) links.forEach((a) => a.classList.toggle('is-active', a.getAttribute('href') === '#' + e.target.id));
      });
    }, { rootMargin: '-45% 0px -50% 0px' });
    $$('main section[id]').forEach((s) => spy.observe(s));
  }

  /* ==================== PESTAÑAS DE HABILIDADES ==================== */
  const tabs = $$('.tab');
  function selectTab(tab) {
    tabs.forEach((x) => {
      const on = x === tab;
      x.classList.toggle('is-active', on);
      x.setAttribute('aria-selected', String(on));
      x.tabIndex = on ? 0 : -1;
      $('#' + x.getAttribute('aria-controls')).hidden = !on;
    });
  }
  tabs.forEach((tab, i) => {
    tab.addEventListener('click', () => selectTab(tab));
    tab.addEventListener('keydown', (e) => {
      if (e.key !== 'ArrowRight' && e.key !== 'ArrowLeft') return;
      const next = tabs[(i + (e.key === 'ArrowRight' ? 1 : tabs.length - 1)) % tabs.length];
      selectTab(next);
      next.focus();
    });
  });

  /* ==================== ANIMACIONES AL ENTRAR / SALIR DEL VIEWPORT ====================
     Dos observadores con márgenes distintos evitan el parpadeo:
     - "enter": muestra el elemento cuando entra en la pantalla.
     - "exit": lo reinicia solo cuando queda fuera del viewport por un margen amplio,
       así el desplazamiento de la propia animación no puede volver a dispararlo. */
  const revealEls = $$('[data-reveal]');
  revealEls.forEach((el) => { if (el.dataset.delay) el.style.setProperty('--d', el.dataset.delay); });

  if (reduceMotion || !('IntersectionObserver' in window)) {
    revealEls.forEach((el) => el.classList.add('is-visible'));
  } else {
    const enter = new IntersectionObserver((entries) => {
      entries.forEach((e) => { if (e.isIntersecting) e.target.classList.add('is-visible'); });
    }, { rootMargin: '0px 0px -8% 0px', threshold: 0 });

    const exit = new IntersectionObserver((entries) => {
      entries.forEach((e) => { if (!e.isIntersecting) e.target.classList.remove('is-visible'); });
    }, { rootMargin: '120px 0px 120px 0px', threshold: 0 });

    revealEls.forEach((el) => { enter.observe(el); exit.observe(el); });
  }

  /* ==================== AÑO ==================== */
  const year = $('#year');
  if (year) year.textContent = new Date().getFullYear();

  /* ==================== INICIO ==================== */
  applyTheme(theme);
  applyLanguage(lang);
})();
