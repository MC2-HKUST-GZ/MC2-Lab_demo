(function () {
  const header = document.querySelector('.leader-header');
  const menuButton = document.getElementById('leaderMenuToggle');

  if (!header || !menuButton) return;

  const closeMenu = () => {
    header.classList.remove('is-menu-open');
    menuButton.setAttribute('aria-expanded', 'false');
  };

  menuButton.addEventListener('click', function () {
    const opening = !header.classList.contains('is-menu-open');
    header.classList.toggle('is-menu-open', opening);
    menuButton.setAttribute('aria-expanded', String(opening));
  });

  document.addEventListener('click', function (event) {
    if (!header.contains(event.target)) {
      closeMenu();
    }
  });

  window.addEventListener('resize', function () {
    if (window.innerWidth > 980) {
      closeMenu();
    }
  });
})();
