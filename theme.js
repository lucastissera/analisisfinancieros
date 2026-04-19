(function () {
  const btn = document.getElementById("btnToggleTheme");
  if (!btn) return;
  btn.addEventListener("click", function () {
    const html = document.documentElement;
    const next = html.getAttribute("data-theme") === "light" ? "dark" : "light";
    html.setAttribute("data-theme", next);
    try {
      localStorage.setItem("analisisFinTema", next);
    } catch (e) {
      /* ignore */
    }
  });
})();
