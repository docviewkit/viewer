const sidebar = document.getElementById("docs-sidebar");
const menuButton = document.querySelector("[data-docs-menu]");
const searchInput = document.getElementById("docs-search-input");

menuButton?.addEventListener("click", () => {
  const open = sidebar?.classList.toggle("open") || false;
  menuButton.setAttribute("aria-expanded", String(open));
});

sidebar?.addEventListener("click", (event) => {
  if (!event.target.closest("a")) return;
  sidebar.classList.remove("open");
  menuButton?.setAttribute("aria-expanded", "false");
});

searchInput?.addEventListener("input", () => {
  const query = searchInput.value.trim().toLowerCase();
  for (const group of sidebar.querySelectorAll(".docs-group")) {
    let matches = 0;
    for (const link of group.querySelectorAll("a")) {
      const match = !query || link.textContent.toLowerCase().includes(query);
      link.hidden = !match;
      if (match) matches += 1;
    }
    group.hidden = matches === 0;
  }
});

document.addEventListener("click", async (event) => {
  const button = event.target.closest("[data-docs-copy]");
  if (!button) return;
  const code = button.closest(".docs-code")?.querySelector("code")?.textContent;
  if (!code) return;
  const previous = button.textContent;
  try {
    await navigator.clipboard.writeText(code);
    button.textContent = "Copied";
  } catch {
    button.textContent = "Copy failed";
  }
  window.setTimeout(() => { button.textContent = previous; }, 1600);
});
