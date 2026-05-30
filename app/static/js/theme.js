/* NEXAS HR — Theme system */
(function () {
  var t = localStorage.getItem('nexas-theme') || 'dark';
  document.documentElement.setAttribute('data-theme', t);
})();

function toggleNexasTheme() {
  var cur = document.documentElement.getAttribute('data-theme') || 'dark';
  var next = cur === 'dark' ? 'light' : 'dark';
  document.documentElement.setAttribute('data-theme', next);
  try { localStorage.setItem('nexas-theme', next); } catch (e) {}
}
