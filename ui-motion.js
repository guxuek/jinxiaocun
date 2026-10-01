/* Visual effects only. No application data, navigation or persistence is changed. */
(function () {
  "use strict";
  var main = document.getElementById("main");
  var nav = document.getElementById("nav");
  if (!main || !nav || !window.MutationObserver || !Element.prototype.animate) return;

  var preference = window.matchMedia("(prefers-reduced-motion: reduce)");
  var ease = "cubic-bezier(.22,1,.36,1)";
  var running = new Set();
  var pendingFrame = 0;
  var activeRoute = "";
  var childrenVisible = [];

  function motionAllowed() { return !preference.matches && !document.hidden; }
  function play(node, frames, options) {
    if (!node || !node.isConnected || !motionAllowed()) return;
    try {
      var effect = node.animate(frames, options);
      running.add(effect);
      function release() { running.delete(effect); }
      effect.onfinish = release;
      effect.oncancel = release;
    } catch (_) { /* The UI works normally when animation support is unavailable. */ }
  }
  function cancelAll() {
    running.forEach(function (effect) { effect.cancel(); });
    running.clear();
  }
  function routeKey() {
    var selected = nav.querySelector(".nav-item.active");
    return selected ? selected.textContent.trim() : "";
  }
  function revealPage() {
    pendingFrame = 0;
    var next = routeKey();
    /* Filtering, pagination and data refreshes keep the page still. */
    if (next === activeRoute) return;
    activeRoute = next;
    cancelAll();
    if (!motionAllowed()) return;
    var sections = Array.prototype.filter.call(main.children, function (node) {
      return !node.classList.contains("busy");
    });
    sections.slice(0, 9).forEach(function (node, i) {
      play(node, [
        { opacity: .25, transform: "translateY(10px)" },
        { opacity: 1, transform: "translateY(0)" }
      ], { duration: 340, delay: Math.min(i * 35, 140), easing: ease });
    });
    /* Only four KPI cards animate; table rows never receive individual effects. */
    var cards = main.querySelectorAll(".home-kpi");
    Array.prototype.forEach.call(cards, function (node, i) {
      play(node, [
        { opacity: .4, transform: "translateY(8px)" },
        { opacity: 1, transform: "translateY(0)" }
      ], { duration: 380, delay: 60 + i * 35, easing: ease });
    });
  }
  function queueReveal() {
    if (!pendingFrame) pendingFrame = requestAnimationFrame(revealPage);
  }
  new MutationObserver(queueReveal).observe(main, { childList: true });

  function revealNavigation() {
    var children = Array.prototype.slice.call(nav.querySelectorAll(".nav-child"));
    var next = children.map(function (node) { return node.textContent.trim(); });
    children.forEach(function (node, i) {
      if (childrenVisible.indexOf(next[i]) >= 0) return;
      play(node, [
        { opacity: .2, transform: "translateX(-5px)" },
        { opacity: 1, transform: "translateX(0)" }
      ], { duration: 230, delay: Math.min(i * 18, 100), easing: ease });
    });
    childrenVisible = next;
  }
  new MutationObserver(revealNavigation).observe(nav, { childList: true });

  /* Delegation also covers dynamically rebuilt buttons and filter chips. */
  document.addEventListener("pointerdown", function (event) {
    if (event.button !== 0 || !motionAllowed()) return;
    var node = event.target.closest(".btn, .chip, .link-btn, .column-filter");
    if (!node || node.disabled) return;
    play(node, [
      { transform: "scale(1)" },
      { transform: "scale(.97)", offset: .4 },
      { transform: "scale(1)" }
    ], { duration: 210, easing: ease });
  }, { passive: true });

  function changedPreference() { if (preference.matches) cancelAll(); }
  if (preference.addEventListener) preference.addEventListener("change", changedPreference);
  else if (preference.addListener) preference.addListener(changedPreference);
  document.addEventListener("visibilitychange", function () { if (document.hidden) cancelAll(); });
  queueReveal();
  revealNavigation();
})();
