/* book-zoom.js — zoom e pan para o book Dom Manuel (deck-stage).
 * Controles:
 *   roda do mouse        → zoom no ponto do cursor (1×–5×)
 *   arrastar             → pan (quando com zoom)
 *   duplo clique         → alterna 2.5× no ponto / 100%
 *   teclas + - 0         → zoom in / out / reset
 *   pinça (touch)        → zoom em tablets/celular
 *   chip "150%"          → clique para resetar
 * Mantido como arquivo separado para sobreviver a re-exports do Claude Design:
 * o book.html precisa apenas do <script src="book-zoom.js"></script> no fim.
 */
(function () {
  'use strict';
  var MIN = 1, MAX = 5, STEP = 1.18;
  var z = 1, tx = 0, ty = 0;
  var stage = null, chip = null, chipTimer = null;
  var dragging = false, dragMoved = false, lastX = 0, lastY = 0;
  var pointers = {}, pinchD0 = 0, pinchZ0 = 1;

  function init() {
    stage = document.querySelector('deck-stage');
    if (!stage) { setTimeout(init, 200); return; }
    stage.style.transformOrigin = '0 0';
    stage.style.willChange = 'transform';

    chip = document.createElement('div');
    chip.style.cssText = 'position:fixed;right:16px;bottom:16px;z-index:99999;' +
      'background:rgba(44,42,36,.82);color:#F2ECD8;font:500 13px/1 Inter,system-ui,sans-serif;' +
      'padding:8px 12px;border-radius:999px;cursor:pointer;letter-spacing:.04em;' +
      'opacity:0;pointer-events:none;transition:opacity .25s;user-select:none;';
    chip.title = 'Clique para voltar a 100%';
    chip.addEventListener('click', function (e) { e.stopPropagation(); reset(); });
    document.body.appendChild(chip);

    window.addEventListener('wheel', onWheel, { passive: false, capture: true });
    window.addEventListener('dblclick', onDbl, true);
    window.addEventListener('keydown', onKey, true);
    window.addEventListener('pointerdown', onDown, true);
    window.addEventListener('pointermove', onMove, true);
    window.addEventListener('pointerup', onUp, true);
    window.addEventListener('pointercancel', onUp, true);
    // suprime o clique que encerraria um arrasto (e o tap-avança quando com zoom)
    window.addEventListener('click', function (e) {
      if (dragMoved) { e.stopPropagation(); e.preventDefault(); dragMoved = false; }
    }, true);
    window.addEventListener('resize', clamp);
  }

  function vw() { return window.innerWidth; }
  function vh() { return window.innerHeight; }

  function apply() {
    clamp();
    stage.style.transform = (z === 1)
      ? ''
      : 'translate(' + tx + 'px,' + ty + 'px) scale(' + z + ')';
    stage.style.cursor = z > 1 ? (dragging ? 'grabbing' : 'grab') : '';
    showChip();
  }

  function clamp() {
    var minX = vw() * (1 - z), minY = vh() * (1 - z);
    tx = Math.min(0, Math.max(minX, tx));
    ty = Math.min(0, Math.max(minY, ty));
  }

  function showChip() {
    if (!chip) return;
    if (z === 1) { chip.style.opacity = '0'; chip.style.pointerEvents = 'none'; return; }
    chip.textContent = Math.round(z * 100) + '%';
    chip.style.opacity = '1';
    chip.style.pointerEvents = 'auto';
    clearTimeout(chipTimer);
    chipTimer = setTimeout(function () { chip.style.opacity = '.35'; }, 1600);
  }

  function zoomAt(cx, cy, nz) {
    nz = Math.min(MAX, Math.max(MIN, nz));
    if (nz === z) return;
    // mantém o ponto (cx,cy) fixo na tela
    tx = cx - (cx - tx) * (nz / z);
    ty = cy - (cy - ty) * (nz / z);
    z = nz;
    apply();
  }

  function reset() { z = 1; tx = 0; ty = 0; apply(); }

  function onWheel(e) {
    if (e.deltaY === 0) return;
    e.preventDefault(); e.stopPropagation();
    zoomAt(e.clientX, e.clientY, z * (e.deltaY < 0 ? STEP : 1 / STEP));
  }

  function onDbl(e) {
    e.preventDefault(); e.stopPropagation();
    if (z === 1) zoomAt(e.clientX, e.clientY, 2.5); else reset();
  }

  function onKey(e) {
    if (e.key === '+' || e.key === '=') { zoomAt(vw() / 2, vh() / 2, z * STEP); }
    else if (e.key === '-' || e.key === '_') { zoomAt(vw() / 2, vh() / 2, z / STEP); }
    else if (e.key === '0') { reset(); }
  }

  function onDown(e) {
    pointers[e.pointerId] = { x: e.clientX, y: e.clientY };
    var ids = Object.keys(pointers);
    if (ids.length === 2) {
      var a = pointers[ids[0]], b = pointers[ids[1]];
      pinchD0 = Math.hypot(a.x - b.x, a.y - b.y);
      pinchZ0 = z;
    } else if (z > 1 && e.isPrimary) {
      dragging = true; dragMoved = false;
      lastX = e.clientX; lastY = e.clientY;
    }
  }

  function onMove(e) {
    if (pointers[e.pointerId]) pointers[e.pointerId] = { x: e.clientX, y: e.clientY };
    var ids = Object.keys(pointers);
    if (ids.length === 2 && pinchD0 > 0) {
      var a = pointers[ids[0]], b = pointers[ids[1]];
      var d = Math.hypot(a.x - b.x, a.y - b.y);
      zoomAt((a.x + b.x) / 2, (a.y + b.y) / 2, pinchZ0 * (d / pinchD0));
      e.preventDefault(); e.stopPropagation();
      return;
    }
    if (dragging) {
      var dx = e.clientX - lastX, dy = e.clientY - lastY;
      if (Math.abs(dx) + Math.abs(dy) > 3) dragMoved = true;
      tx += dx; ty += dy;
      lastX = e.clientX; lastY = e.clientY;
      apply();
      e.preventDefault(); e.stopPropagation();
    }
  }

  function onUp(e) {
    delete pointers[e.pointerId];
    if (Object.keys(pointers).length < 2) pinchD0 = 0;
    if (dragging && e.isPrimary) { dragging = false; apply(); }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else { init(); }
})();
