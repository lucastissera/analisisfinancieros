import { USUARIOS_PERMITIDOS } from "./usuarios-permitidos.js";

const AUTH_KEY = "analisisFinAuthV1";
/** Última actividad (timestamp ms) para expiración de sesión. */
const ACTIVITY_KEY = "analisisFinLastActivityV1";

/** Duración máxima de inactividad antes de pedir credenciales de nuevo. */
const SESSION_MS = 30 * 60 * 1000;

/** Cada cuánto se comprueba si venció la sesión (mientras la app está abierta). */
const CHECK_INTERVAL_MS = 15 * 1000;

/** Mínimo entre actualizaciones de “última actividad” por eventos del usuario. */
const ACTIVITY_THROTTLE_MS = 10 * 1000;

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

function touchActivity() {
  sessionStorage.setItem(ACTIVITY_KEY, String(Date.now()));
}

function isSessionExpired() {
  const raw = sessionStorage.getItem(ACTIVITY_KEY);
  if (!raw) return true;
  const last = Number(raw);
  if (!Number.isFinite(last)) return true;
  return Date.now() - last > SESSION_MS;
}

function clearSession() {
  sessionStorage.removeItem(AUTH_KEY);
  sessionStorage.removeItem(ACTIVITY_KEY);
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

function mostrarLogin() {
  const login = $("view-login");
  const app = $("app-shell");
  if (!login || !app) return;
  login.hidden = false;
  app.hidden = true;
}

let expiryTimerId = null;
let lastThrottleMark = 0;

function onUserActivity() {
  if (sessionStorage.getItem(AUTH_KEY) !== "1") return;
  const now = Date.now();
  if (now - lastThrottleMark < ACTIVITY_THROTTLE_MS) return;
  lastThrottleMark = now;
  touchActivity();
}

function cerrarSesionPorExpiracion() {
  clearSession();
  detenerMonitorSesion();
  mostrarLogin();
  const exp = $("loginSesionExpirada");
  const err = $("loginErr");
  if (exp) exp.hidden = false;
  if (err) err.hidden = true;
}

function detenerMonitorSesion() {
  if (expiryTimerId != null) {
    clearInterval(expiryTimerId);
    expiryTimerId = null;
  }
  document.removeEventListener("click", onUserActivity, true);
  document.removeEventListener("keydown", onUserActivity, true);
  document.removeEventListener("pointerdown", onUserActivity, true);
  document.removeEventListener("visibilitychange", onVisibilityChange);
}

function onVisibilityChange() {
  if (document.visibilityState !== "visible") return;
  if (sessionStorage.getItem(AUTH_KEY) !== "1") return;
  if (isSessionExpired()) {
    cerrarSesionPorExpiracion();
  }
}

function iniciarMonitorSesion() {
  detenerMonitorSesion();
  touchActivity();
  lastThrottleMark = Date.now();
  expiryTimerId = setInterval(() => {
    if (sessionStorage.getItem(AUTH_KEY) !== "1") {
      detenerMonitorSesion();
      return;
    }
    if (isSessionExpired()) {
      cerrarSesionPorExpiracion();
    }
  }, CHECK_INTERVAL_MS);
  document.addEventListener("click", onUserActivity, true);
  document.addEventListener("keydown", onUserActivity, true);
  document.addEventListener("pointerdown", onUserActivity, true);
  document.addEventListener("visibilitychange", onVisibilityChange);
}

function intentarLogin() {
  const err = $("loginErr");
  const exp = $("loginSesionExpirada");
  err.hidden = true;
  if (exp) exp.hidden = true;
  if (credencialesValidas($("loginUsuario").value, $("loginClave").value)) {
    sessionStorage.setItem(AUTH_KEY, "1");
    touchActivity();
    mostrarApp();
    iniciarMonitorSesion();
  } else {
    err.hidden = false;
  }
}

function init() {
  if (sessionStorage.getItem(AUTH_KEY) === "1") {
    if (isSessionExpired()) {
      clearSession();
      mostrarLogin();
    } else {
      touchActivity();
      mostrarApp();
      iniciarMonitorSesion();
    }
  }

  const u = $("loginUsuario");
  const c = $("loginClave");
  const btn = $("btnLogin");

  function ocultarError() {
    $("loginErr").hidden = true;
    const exp = $("loginSesionExpirada");
    if (exp) exp.hidden = true;
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
