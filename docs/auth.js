"use strict";

// ---------------------------------------------------------------------
// Hardcoded single-account login gate. This only controls who sees the
// UI in this browser tab — it does not (and cannot, on a static site)
// hide the underlying source files from someone who requests them
// directly. See login-page-feasibility-review.md at the repo root.
// ---------------------------------------------------------------------

const AUTH_USERNAME = "suraj.singh@gmail.com";
const AUTH_PASSWORD = "Suraj@Singh";
const AUTH_SESSION_KEY = "gst_scrutiny_auth_ok";

function isLoggedIn() {
  return sessionStorage.getItem(AUTH_SESSION_KEY) === "1";
}

function showApp() {
  document.getElementById("login-overlay").hidden = true;
  document.getElementById("app-shell").hidden = false;
}

function showLogin() {
  document.getElementById("app-shell").hidden = true;
  document.getElementById("login-overlay").hidden = false;
}

function switchView(view) {
  document.querySelectorAll(".app-view").forEach((el) => {
    el.hidden = el.id !== `view-${view}`;
  });
  document.querySelectorAll(".app-nav-item[data-view]").forEach((el) => {
    el.classList.toggle("active", el.dataset.view === view);
  });
}

function handleLoginSubmit(event) {
  event.preventDefault();
  const username = document.getElementById("login-username").value.trim();
  const password = document.getElementById("login-password").value;
  const errorEl = document.getElementById("login-error");

  const matches =
    username.toLowerCase() === AUTH_USERNAME.toLowerCase() &&
    password === AUTH_PASSWORD;

  if (!matches) {
    errorEl.textContent = "Incorrect email or password.";
    errorEl.hidden = false;
    return;
  }

  errorEl.hidden = true;
  sessionStorage.setItem(AUTH_SESSION_KEY, "1");
  showApp();
  document.dispatchEvent(new Event("auth:ok"));
}

function handleLogout() {
  sessionStorage.removeItem(AUTH_SESSION_KEY);
  document.getElementById("login-form").reset();
  document.getElementById("login-error").hidden = true;
  switchView("welcome");
  showLogin();
}

document.addEventListener("DOMContentLoaded", () => {
  document.getElementById("login-form").addEventListener("submit", handleLoginSubmit);
  document.getElementById("logout-btn").addEventListener("click", handleLogout);
  document.getElementById("welcome-go-btn").addEventListener("click", () => switchView("scrutinize"));

  document.querySelectorAll(".app-nav-item[data-view]").forEach((el) => {
    el.addEventListener("click", () => switchView(el.dataset.view));
  });

  if (isLoggedIn()) {
    showApp();
    document.dispatchEvent(new Event("auth:ok"));
  } else {
    showLogin();
  }
});
