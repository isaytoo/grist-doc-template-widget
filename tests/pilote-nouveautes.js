(async function () {
  var R = {}; var attendre = function (ms) { return new Promise(function (r) { setTimeout(r, ms); }); };
  var ok = function (nom, cond, detail) { R[(cond ? 'OK  ' : 'KO  ') + nom] = detail === undefined ? '' : String(detail); };
  try {
    document.getElementById('main-content').classList.remove('hidden');
    initEditor(); await attendre(300);
    var ed = editorInstance;
    var T = '<table style="width:100%"><tbody><tr><td>A1</td><td>B1</td></tr><tr><td>A2</td><td>B2</td></tr></tbody></table>';
    var fermer = function () { var b = document.querySelector('#border-cancel, #cell-border-cancel'); if (b) b.click(); };
    // ce que l'export PDF dessinerait pour chaque cellule
    var bordsPdf = async function (html) {
      var vus = []; var garde = window.html2canvas;
      window.html2canvas = async function (el) { el.querySelectorAll('td').forEach(function (td) { var c = getComputedStyle(td); vus.push(c.borderTopStyle === 'none' ? 'aucune' : c.borderTopWidth + ' ' + c.borderTopStyle + ' ' + rgbToHex(c.borderTopColor)); }); var cv = document.createElement('canvas'); cv.width = 10; cv.height = 10; return cv; };
      await renderHtmlToPdfPages(html, { addPage: function () {}, addImage: function () {} }, 210, 297, 'a4');
      window.html2canvas = garde; return vus;
    };

    // ── tableau entier invisible ──
    ed.value = T; await attendre(50);
    ed.s.setCursorIn(ed.editor.querySelector('td'));
    ed.o.controls.tableborder.exec(ed);
    ok('réglage tableau : case décochée au départ', document.querySelector('#border-invisible').checked === false);
    document.querySelector('#border-invisible').checked = true; document.querySelector('#border-apply').click(); await attendre(50);
    ed.s.setCursorIn(ed.editor.querySelector('td'));
    ed.o.controls.tableborder.exec(ed);
    ok('réglage tableau rouvert : « Bordures invisibles » reste coché', document.querySelector('#border-invisible').checked === true); fermer();
    var b = await bordsPdf(ed.value);
    ok('PDF : aucune bordure sur les quatre cellules', b.join('|') === 'aucune|aucune|aucune|aucune', b.join('|'));
    await attendre(2400);
    ok('sauvegarde automatique déclenchée par le réglage', /border-style:\s*none/.test(window.__options.editorDraft || ''));

    // ── on les remet, en rouge pointillé ──
    ed.s.setCursorIn(ed.editor.querySelector('td'));
    ed.o.controls.tableborder.exec(ed);
    document.querySelector('#border-invisible').checked = false;
    document.querySelector('#border-color').value = '#ff0000'; document.querySelector('#border-width').value = '2px'; document.querySelector('#border-style').value = 'dashed';
    document.querySelector('#border-apply').click(); await attendre(50);
    b = await bordsPdf(ed.value);
    ok('PDF : bordure choisie conservée', b[0] === '2px dashed #ff0000', b[0]);
    ed.s.setCursorIn(ed.editor.querySelector('td')); ed.o.controls.tableborder.exec(ed);
    ok('réglage rouvert : couleur, épaisseur et style retrouvés', document.querySelector('#border-color').value === '#ff0000' && document.querySelector('#border-width').value === '2px' && document.querySelector('#border-style').value === 'dashed', document.querySelector('#border-color').value + ' ' + document.querySelector('#border-width').value + ' ' + document.querySelector('#border-style').value); fermer();

    // ── une seule cellule invisible ──
    ed.value = T; await attendre(50);
    ed.s.setCursorIn(ed.editor.querySelector('td'));
    ed.o.controls.cellborder.exec(ed);
    document.querySelector('#cell-border-invisible').checked = true; document.querySelector('#cell-border-apply').click(); await attendre(50);
    ed.s.setCursorIn(ed.editor.querySelector('td')); ed.o.controls.cellborder.exec(ed);
    ok('réglage cellule rouvert : case cochée', document.querySelector('#cell-border-invisible').checked === true); fermer();
    var tmp = document.createElement('div'); tmp.innerHTML = ed.value;
    ok('seule la cellule visée porte le réglage', hasInlineBorder(tmp.querySelectorAll('td')[0]) && !hasInlineBorder(tmp.querySelectorAll('td')[1]));

    // ── tableau sans aucun réglage : comme avant ──
    b = await bordsPdf(T);
    ok('PDF : tableau ordinaire toujours bordé de noir', b.join('|') === '1px solid #000000|1px solid #000000|1px solid #000000|1px solid #000000', b[0]);

    // ── mise en forme des variables ──
    var puce = '<span style="background:#f3e8ff;color:#7c3aed;padding:2px 6px;border-radius:4px;font-weight:600;" contenteditable="false">{{Nom}}</span>&nbsp;';
    var styleDuNom = function (html, prop) { var d = document.createElement('div'); d.style.cssText = 'font-family:"Times New Roman",Times,serif;font-size:14px'; d.innerHTML = html; document.body.appendChild(d);
      var w = document.createTreeWalker(d, NodeFilter.SHOW_TEXT), n, f = '?'; while ((n = w.nextNode())) if (n.nodeValue.indexOf('Durand') !== -1) f = getComputedStyle(n.parentNode)[prop]; d.remove(); return f; };
    var toutSelectionner = function () { var rg = document.createRange(); rg.selectNodeContents(ed.editor.querySelector('p')); var s = window.getSelection(); s.removeAllRanges(); s.addRange(rg); };
    var essai = async function (cmd, val) { ed.value = '<p>Bonjour ' + puce + 'et bienvenue.</p>'; await attendre(50); toutSelectionner(); ed.execCommand(cmd, false, val); await attendre(50); return ed.value; };

    var m = await essai('fontName', 'Arial,Helvetica,sans-serif');
    ok('police : le champ fusionné la prend (PDF)', /Arial/.test(styleDuNom(await resolveTemplateAsync(m, { Nom: 'Durand' }, true), 'fontFamily')), styleDuNom(await resolveTemplateAsync(m, { Nom: 'Durand' }, true), 'fontFamily'));
    ok('police : aussi dans l’aperçu', /Arial/.test(styleDuNom(resolveTemplate(m, { Nom: 'Durand' }, false), 'fontFamily')));
    ok('la puce est de nouveau verrouillée', ed.editor.querySelector('span[style*="border-radius"]').getAttribute('contenteditable') === 'false');
    ok('la puce garde son habillage dans l’éditeur', /#f3e8ff|243, 232, 255/.test(ed.editor.querySelector('span[style*="border-radius"]').getAttribute('style')));
    ok('aucun habillage de puce ne fuit dans le PDF', !/border-radius|f3e8ff|243, 232/.test(await resolveTemplateAsync(m, { Nom: 'Durand' }, true)));
    m = await essai('fontSize', '24px');
    ok('taille : le champ la prend', styleDuNom(await resolveTemplateAsync(m, { Nom: 'Durand' }, true), 'fontSize') === '24px', styleDuNom(await resolveTemplateAsync(m, { Nom: 'Durand' }, true), 'fontSize'));
    m = await essai('bold');
    ok('gras : le champ le prend', parseInt(styleDuNom(await resolveTemplateAsync(m, { Nom: 'Durand' }, true), 'fontWeight'), 10) >= 700, styleDuNom(await resolveTemplateAsync(m, { Nom: 'Durand' }, true), 'fontWeight'));
    m = await essai('italic');
    ok('italique : le champ le prend', styleDuNom(await resolveTemplateAsync(m, { Nom: 'Durand' }, true), 'fontStyle') === 'italic');
    // sans mise en forme : ni gras ni couleur hérités de la puce
    var brut = await resolveTemplateAsync('<p>Bonjour ' + puce + '</p>', { Nom: 'Durand' }, true);
    ok('sans mise en forme : pas de gras ni de violet hérités de la puce', parseInt(styleDuNom(brut, 'fontWeight'), 10) === 400 && !/124, 58, 237/.test(styleDuNom(brut, 'color')), styleDuNom(brut, 'fontWeight') + ' / ' + styleDuNom(brut, 'color'));
    // ── version affichée ──
    ok('version affichée près du titre', /^v\d+$/.test(document.getElementById('widget-version-titre').textContent), document.getElementById('widget-version-titre').textContent);
    ok('version affichée en pied de page (placé après le script)', document.getElementById('widget-version').textContent === WIDGET_VERSION, document.getElementById('widget-version').textContent);
    ok('elle correspond au ?v= de index.html', __VERSION_INDEX__ === WIDGET_VERSION.slice(1), __VERSION_INDEX__ + ' / ' + WIDGET_VERSION);
  } catch (e) { R['KO  erreur'] = String(e && e.stack || e).slice(0, 500); }
  document.getElementById('out').textContent = 'RES' + JSON.stringify(R) + 'FIN';
})();
