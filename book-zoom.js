/* book-zoom.js — zoom e pan para o book Dom Manuel (deck-stage).
 * Barra de controle fixa (canto inferior direito):  −  100%  +
 *   −/+                  → zoom out/in centrado na tela
 *   clique no percentual → volta a 100%
 * Extras (não obrigatórios):
 *   arrastar → pan (quando ampliado) · teclas + - 0 · pinça em touch
 *   (roda do mouse e duplo clique ficam LIVRES para o deck — decisão 04/07/2026)
 * Arquivo separado para sobreviver a re-exports do Claude Design:
 * o book.html precisa apenas do <script src="book-zoom.js"></script> no fim.
 */
(function () {
  'use strict';
  var MIN = 1, MAX = 5, STEP = 1.25;
  var z = 1, tx = 0, ty = 0;
  var stage = null, bar = null, lbl = null;
  var dragging = false, dragMoved = false, lastX = 0, lastY = 0;
  var pointers = {}, pinchD0 = 0, pinchZ0 = 1;

  function init() {
    stage = document.querySelector('deck-stage');
    if (!stage) { setTimeout(init, 200); return; }
    stage.style.transformOrigin = '0 0';
    // o deck marca .stage/.canvas com will-change:transform, o que congela a
    // rasterização no tamanho original e deixa o zoom borrado; neutraliza:
    try {
      if (stage.shadowRoot) {
        var st = document.createElement('style');
        st.textContent = '.stage,.canvas{will-change:auto !important;}';
        stage.shadowRoot.appendChild(st);
      }
    } catch (err) {}
    buildBar();

    window.addEventListener('keydown', onKey, true);
    window.addEventListener('pointerdown', onDown, true);
    window.addEventListener('pointermove', onMove, true);
    window.addEventListener('pointerup', onUp, true);
    window.addEventListener('pointercancel', onUp, true);
    window.addEventListener('click', function (e) {
      if (dragMoved) { e.stopPropagation(); e.preventDefault(); dragMoved = false; }
    }, true);
    // com zoom ativo, impede o drag nativo de imagens e a seleção de texto,
    // que interrompem o arrasto (pan)
    window.addEventListener('dragstart', function (e) {
      if (z > 1) { e.preventDefault(); e.stopPropagation(); }
    }, true);
    window.addEventListener('selectstart', function (e) {
      if (z > 1) { e.preventDefault(); }
    }, true);
    window.addEventListener('resize', function () { clamp(); render(); });
  }

  function buildBar() {
    bar = document.createElement('div');
    bar.style.cssText =
      'position:fixed;right:18px;bottom:18px;z-index:99999;display:flex;align-items:center;gap:2px;' +
      'background:rgba(44,42,36,.88);border:1px solid rgba(221,208,180,.35);border-radius:999px;' +
      'padding:4px;box-shadow:0 4px 18px rgba(0,0,0,.35);user-select:none;' +
      'opacity:.55;transition:opacity .2s;';
    bar.addEventListener('mouseenter', function () { bar.style.opacity = '1'; });
    bar.addEventListener('mouseleave', function () { bar.style.opacity = z > 1 ? '.9' : '.55'; });
    // impede que cliques na barra cheguem ao deck
    ['pointerdown', 'pointerup', 'click', 'dblclick', 'wheel', 'touchstart'].forEach(function (ev) {
      bar.addEventListener(ev, function (e) { e.stopPropagation(); }, false);
    });

    function btn(text, title, fn, wide) {
      var b = document.createElement('button');
      b.type = 'button';
      b.textContent = text;
      b.title = title;
      b.style.cssText =
        'appearance:none;border:0;background:transparent;color:#F2ECD8;cursor:pointer;' +
        'font:600 ' + (wide ? '13' : '20') + 'px/1 Inter,system-ui,sans-serif;letter-spacing:.03em;' +
        'min-width:' + (wide ? '64' : '40') + 'px;height:40px;border-radius:999px;padding:0 10px;' +
        'display:flex;align-items:center;justify-content:center;';
      b.addEventListener('mouseenter', function () { b.style.background = 'rgba(221,208,180,.18)'; });
      b.addEventListener('mouseleave', function () { b.style.background = 'transparent'; });
      b.addEventListener('click', function (e) { e.stopPropagation(); e.preventDefault(); fn(); });
      bar.appendChild(b);
      return b;
    }

    btn('−', 'Diminuir zoom', function () { zoomAt(cx(), cy(), z / STEP); });
    lbl = btn('100%', 'Voltar a 100%', reset, true);
    lbl.style.color = '#DDD0B4';
    btn('+', 'Aumentar zoom', function () { zoomAt(cx(), cy(), z * STEP); });

    document.body.appendChild(bar);
  }

  function cx() { return window.innerWidth / 2; }
  function cy() { return window.innerHeight / 2; }

  function paint() {
    // transform:scale no host — não altera layout, então o deck não "refita".
    // A nitidez vem da neutralização do will-change (ver init), que permite ao
    // navegador re-rasterizar o conteúdo na escala ampliada.
    stage.style.transform = (z === 1) ? '' : 'translate(' + tx + 'px,' + ty + 'px) scale(' + z + ')';
  }

  function render() {
    paint();
    clampToSlide();
    stage.style.cursor = z > 1 ? (dragging ? 'grabbing' : 'grab') : '';
    if (lbl) lbl.textContent = Math.round(z * 100) + '%';
    if (bar && !bar.matches(':hover')) bar.style.opacity = z > 1 ? '.9' : '.55';
  }

  function clamp() {
    // limite bruto (host inteiro) — refinado depois por clampToSlide()
    var minX = window.innerWidth * (1 - z), minY = window.innerHeight * (1 - z);
    tx = Math.min(0, Math.max(minX, tx));
    ty = Math.min(0, Math.max(minY, ty));
  }

  function slideRect() {
    // retângulo do slide na tela (área útil, sem o letterbox preto do deck)
    try {
      var inner = stage.shadowRoot && stage.shadowRoot.querySelector('.stage');
      if (inner) return inner.getBoundingClientRect();
    } catch (err) {}
    return stage.getBoundingClientRect();
  }

  function clampToSlide() {
    // ajusta tx/ty para que a janela visível fique dentro do slide
    if (z === 1) return;
    var r = slideRect(), W = window.innerWidth, H = window.innerHeight;
    var dx = 0, dy = 0;
    if (r.width <= W) dx = (W - r.width) / 2 - r.left;        // centra
    else if (r.left > 0) dx = -r.left;                         // sobra à esquerda
    else if (r.right < W) dx = W - r.right;                    // sobra à direita
    if (r.height <= H) dy = (H - r.height) / 2 - r.top;
    else if (r.top > 0) dy = -r.top;
    else if (r.bottom < H) dy = H - r.bottom;
    if (dx || dy) { tx += dx; ty += dy; paint(); }
  }

  function zoomAt(px, py, nz) {
    nz = Math.min(MAX, Math.max(MIN, nz));
    if (nz === z) return;
    tx = px - (px - tx) * (nz / z);
    ty = py - (py - ty) * (nz / z);
    z = nz;
    clamp(); render();
  }

  function reset() { z = 1; tx = 0; ty = 0; render(); }

  function onKey(e) {
    if (e.key === '+' || e.key === '=') zoomAt(cx(), cy(), z * STEP);
    else if (e.key === '-' || e.key === '_') zoomAt(cx(), cy(), z / STEP);
    else if (e.key === '0') reset();
  }

  function onDown(e) {
    if (bar && bar.contains(e.target)) return;
    pointers[e.pointerId] = { x: e.clientX, y: e.clientY };
    var ids = Object.keys(pointers);
    if (ids.length === 2) {
      var a = pointers[ids[0]], b = pointers[ids[1]];
      pinchD0 = Math.hypot(a.x - b.x, a.y - b.y);
      pinchZ0 = z;
    } else if (z > 1 && e.isPrimary) {
      dragging = true; dragMoved = false;
      lastX = e.clientX; lastY = e.clientY;
      e.preventDefault(); // bloqueia drag nativo/seleção já no início do gesto
      try { e.target.setPointerCapture && e.target.setPointerCapture(e.pointerId); } catch (err) {}
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
      clamp(); render();
      e.preventDefault(); e.stopPropagation();
    }
  }

  function onUp(e) {
    delete pointers[e.pointerId];
    if (Object.keys(pointers).length < 2) pinchD0 = 0;
    if (dragging && e.isPrimary) { dragging = false; render(); }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else { init(); }
})();
