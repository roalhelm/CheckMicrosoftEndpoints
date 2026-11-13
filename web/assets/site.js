// Site chrome shared with letmetrythis.dev: mobile nav, theme toggle, footer year.
// Dark is the default; light is stored only after an explicit choice (same key as the blog).
(function () {
  var root = document.documentElement;
  var toggle = document.querySelector('.nav-toggle');
  var nav = document.getElementById('main-nav');
  if (toggle && nav) {
    toggle.addEventListener('click', function () {
      var open = nav.classList.toggle('open');
      toggle.setAttribute('aria-expanded', String(open));
    });
    Array.prototype.forEach.call(nav.querySelectorAll('a'), function (link) {
      link.addEventListener('click', function () {
        nav.classList.remove('open');
        toggle.setAttribute('aria-expanded', 'false');
      });
    });
  }

  var header = document.querySelector('.site-header');
  if (header) {
    var btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'theme-toggle';
    var render = function () {
      var light = root.dataset.theme === 'light';
      btn.textContent = light ? '☀' : '☾';
      var label = light ? 'Switch to dark mode' : 'Switch to light mode';
      btn.setAttribute('aria-label', label);
      btn.title = label;
      var meta = document.querySelector('meta[name="theme-color"]');
      if (meta) meta.setAttribute('content', light ? '#f4f8fb' : '#0b1120');
    };
    btn.addEventListener('click', function () {
      var light = root.dataset.theme !== 'light';
      if (light) root.dataset.theme = 'light'; else delete root.dataset.theme;
      try { localStorage.setItem('theme', light ? 'light' : 'dark'); } catch (e) {}
      render();
    });
    render();
    header.appendChild(btn);
  }

  var year = document.getElementById('year');
  if (year) year.textContent = new Date().getFullYear();
})();
