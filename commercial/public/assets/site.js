import { initializeI18n, t } from "./i18n.js?v=20261002-3";

const storageKey = "docviewkit-theme";

function initialTheme() {
  const stored = localStorage.getItem(storageKey);
  if (stored === "light" || stored === "dark") return stored;
  return window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
}

function applyTheme(theme) {
  document.documentElement.dataset.theme = theme;
  document.querySelectorAll("[data-theme-toggle]").forEach((button) => {
    button.setAttribute("aria-label", t(theme === "dark" ? "切换到浅色主题" : "切换到深色主题"));
  });
}

initializeI18n();
applyTheme(initialTheme());

document.addEventListener("click", async (event) => {
  const themeButton = event.target.closest("[data-theme-toggle]");
  if (themeButton) {
    const theme = document.documentElement.dataset.theme === "dark" ? "light" : "dark";
    localStorage.setItem(storageKey, theme);
    applyTheme(theme);
    return;
  }

  const menuButton = event.target.closest("[data-menu-toggle]");
  if (menuButton) {
    const menu = document.getElementById(menuButton.getAttribute("aria-controls"));
    const open = menu?.classList.toggle("open") || false;
    menuButton.setAttribute("aria-expanded", String(open));
    return;
  }

  const navLink = event.target.closest("#primary-nav a");
  if (navLink) {
    document.getElementById("primary-nav")?.classList.remove("open");
    document.querySelector("[data-menu-toggle]")?.setAttribute("aria-expanded", "false");
  }

  const copyButton = event.target.closest("[data-copy-target]");
  if (copyButton) {
    const target = document.getElementById(copyButton.dataset.copyTarget);
    if (!target) return;
    const previous = copyButton.textContent;
    try {
      await navigator.clipboard.writeText(target.textContent);
      copyButton.textContent = t("已复制");
    } catch {
      copyButton.textContent = t("复制失败");
    }
    window.setTimeout(() => { copyButton.textContent = previous; }, 1800);
  }
});

document.querySelectorAll("[data-year]").forEach((node) => {
  node.textContent = String(new Date().getFullYear());
});
