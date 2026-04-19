import { USUARIOS_PERMITIDOS } from "./usuarios-permitidos.js";

const AUTH_KEY = "analisisFinAuthV1";

function $(id) {
  return document.getElementById(id);
}

function credencialesValidas(usuarioRaw, claveRaw) {
  const usuario = String(usuarioRaw ?? "").trim();
  const clave = String(claveRaw ?? "");
  return USUARIOS_PERMITIDOS.some(
    (row) =>
      String(row.usuario ?? "").trim() === usuario &&
      String(row.clave ?? "") === clave
  );
}

function actualizarBotonLogin() {
  const u = $("loginUsuario").value.trim();
  const p = $("loginClave").value;
  $("btnLogin").disabled = !u || !p;
}

function mostrarApp() {
  const login = $("view-login");
  const app = $("app-shell");
  if (!login || !app) return;
  login.hidden = true;
  app.hidden = false;
}

function intentarLogin() {
  const err = $("loginErr");
  err.hidden = true;
  if (credencialesValidas($("loginUsuario").value, $("loginClave").value)) {
    sessionStorage.setItem(AUTH_KEY, "1");
    mostrarApp();
  } else {
    err.hidden = false;
  }
}

function init() {
  if (sessionStorage.getItem(AUTH_KEY) === "1") {
    mostrarApp();
  }

  const u = $("loginUsuario");
  const c = $("loginClave");
  const btn = $("btnLogin");

  function ocultarError() {
    $("loginErr").hidden = true;
  }

  u.addEventListener("input", () => {
    actualizarBotonLogin();
    ocultarError();
  });
  c.addEventListener("input", () => {
    actualizarBotonLogin();
    ocultarError();
  });

  btn.addEventListener("click", intentarLogin);
  c.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !btn.disabled) intentarLogin();
  });

  $("btnUsuarioNuevo").addEventListener("click", () => {
    const text =
      "Buen día! quisiera generar mi usuario en el sistema de análisis financieros";
    const url = `https://wa.me/5493513132914?text=${encodeURIComponent(text)}`;
    window.open(url, "_blank", "noopener,noreferrer");
  });

  actualizarBotonLogin();
}

init();
