/**
 * Acceso de prueba (revertir al integrar auth real).
 */
const LOGIN_USER = "admin";
const LOGIN_PASS = "admin";

const viewLogin = document.getElementById("view-login");
const viewApp = document.getElementById("view-app");
const loginErr = document.getElementById("loginErr");
const loginForm = document.getElementById("loginForm");

function intentarLogin() {
  const u = document.getElementById("loginUser").value.trim();
  const p = document.getElementById("loginPass").value;
  if (u === LOGIN_USER && p === LOGIN_PASS) {
    loginErr.hidden = true;
    viewLogin.hidden = true;
    viewApp.hidden = false;
    document.title = "Análisis de FCI";
  } else {
    loginErr.textContent = "Usuario o contraseña incorrectos.";
    loginErr.hidden = false;
  }
}

loginForm.addEventListener("submit", (e) => {
  e.preventDefault();
  intentarLogin();
});
