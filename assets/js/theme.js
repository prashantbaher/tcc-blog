/**
 * Theme toggle + sidebar + mobile menu behavior
 */
document.addEventListener("DOMContentLoaded", function () {
  const root = document.documentElement;
  const toggleBtn = document.getElementById("theme-toggle");
  const iconMoon = "&#xE708;"; // moon
  const iconSun = "&#xE706;"; // sun

  function setThemeButton(theme) {
    if (!toggleBtn) return;
    const glyph = theme === "dark" ? iconSun : iconMoon;
    const label = theme === "dark" ? "Switch to light mode" : "Switch to dark mode";
    toggleBtn.innerHTML = '<span class="ms-icon" aria-hidden="true">' + glyph + "</span>";
    toggleBtn.setAttribute("aria-label", label);
  }

  function applyTheme(theme) {
    root.setAttribute("data-theme", theme);
    root.classList.toggle("theme-dark", theme === "dark");
    setThemeButton(theme);
  }

  const storedTheme = localStorage.getItem("jekyll-doks-theme");
  if (storedTheme) {
    applyTheme(storedTheme);
  } else {
    const prefersDark =
      window.matchMedia &&
      window.matchMedia("(prefers-color-scheme: dark)").matches;
    applyTheme(prefersDark ? "dark" : "light");
  }

  if (toggleBtn) {
    toggleBtn.addEventListener("click", function () {
      const currentTheme = root.getAttribute("data-theme") || "light";
      const nextTheme = currentTheme === "light" ? "dark" : "light";
      applyTheme(nextTheme);
      localStorage.setItem("jekyll-doks-theme", nextTheme);
    });
  }

  if (window.matchMedia) {
    window
      .matchMedia("(prefers-color-scheme: dark)")
      .addEventListener("change", function (e) {
        if (!localStorage.getItem("jekyll-doks-theme")) {
          applyTheme(e.matches ? "dark" : "light");
        }
      });
  }

  const mobileMenuToggle = document.getElementById("mobile-menu-toggle");
  const mainNav = document.getElementById("main-nav");

  function setMenuToggle(open) {
    if (!mobileMenuToggle) return;
    const glyph = open ? "&#xE711;" : "&#xE700;"; // close / menu
    mobileMenuToggle.innerHTML = '<span class="ms-icon" aria-hidden="true">' + glyph + "</span>";
    mobileMenuToggle.setAttribute("aria-expanded", String(open));
  }

  if (mobileMenuToggle && mainNav) {
    setMenuToggle(false);

    mobileMenuToggle.addEventListener("click", function (e) {
      e.stopPropagation();
      mainNav.classList.toggle("active");
      setMenuToggle(mainNav.classList.contains("active"));
    });

    document.addEventListener("click", function (event) {
      if (!event.target.closest(".site-header")) {
        mainNav.classList.remove("active");
        setMenuToggle(false);
      }
    });

    const navLinks = mainNav.querySelectorAll("a");
    navLinks.forEach(function (link) {
      link.addEventListener("click", function () {
        mainNav.classList.remove("active");
        setMenuToggle(false);
      });
    });
  }

  const sidebarSections = document.querySelectorAll(".sidebar-section");
  sidebarSections.forEach(function (section) {
    const button = section.querySelector(".section-toggle");
    const list = section.querySelector(".section-links");
    if (!button || !list) return;

    const hasActiveLink = list.querySelector("a.active");
    const expanded = Boolean(hasActiveLink);
    list.style.display = expanded ? "block" : "none";
    button.setAttribute("aria-expanded", String(expanded));

    button.addEventListener("click", function () {
      const isVisible = list.style.display !== "none";
      list.style.display = isVisible ? "none" : "block";
      button.setAttribute("aria-expanded", String(!isVisible));
    });
  });

  const backToTopBtn = document.querySelector(".back-to-top");
  if (backToTopBtn) {
    function toggleBackToTop() {
      backToTopBtn.style.display = window.scrollY > 300 ? "inline-block" : "none";
    }

    window.addEventListener("scroll", toggleBackToTop, { passive: true });
    toggleBackToTop();

    backToTopBtn.addEventListener("click", function () {
      window.scrollTo({ top: 0, behavior: "smooth" });
    });
  }
});
