// Mêmes entrées passées à l'ancienne et à la nouvelle version : pour des modèles
// sans mise en forme sur les variables, la sortie doit être identique.
(async function () {
  var R = {}; var attendre = function (ms) { return new Promise(function (r) { setTimeout(r, ms); }); };
  try {
    var puce = function (n) { return '<span style="background:#f3e8ff;color:#7c3aed;padding:2px 6px;border-radius:4px;font-weight:600;" contenteditable="false">{{' + n + '}}</span>&nbsp;'; };
    var img = '<span style="background:#dcfce7;color:#166534;padding:2px 6px;border-radius:4px;">{{IMG:Photo:120}}</span>';
    var cas = {
      simple: '<p>Bonjour ' + puce('Nom') + ', né le ' + puce('Date') + '</p>',
      nu: '<p>Bonjour {{Nom}} et {{Inconnu}}</p>',
      vide: '<p>Ville : ' + puce('Ville') + '</p>',
      tableau: '<table style="width:100%"><tbody><tr><td>' + puce('Nom') + '</td><td style="border: 2px dashed rgb(255, 0, 0);">{{Montant}}</td></tr></tbody></table>',
      image: '<p>' + img + '</p>',
      saut: '<p>a</p><div class="page-break-marker" contenteditable="false">— saut —</div><p>' + puce('Nom') + '</p>',
      special: '<p>' + puce('Nom') + ' &amp; <b>{{Nom}}</b></p>'
    };
    var rec = { Nom: 'Durand <b>&', Date: 1762819200, Ville: '', Montant: 1234.5, Photo: '' };
    for (var k in cas) {
      R[k + ' | aperçu'] = resolveTemplate(cas[k], rec, false);
      R[k + ' | pdf'] = resolveTemplate(cas[k], rec, true);
      R[k + ' | pdf async'] = await resolveTemplateAsync(cas[k], rec, true);
      R[k + ' | aperçu async'] = await resolveTemplateAsync(cas[k], rec, false);
    }
  } catch (e) { R.erreur = String(e && e.stack || e).slice(0, 400); }
  document.getElementById('out').textContent = 'RES' + JSON.stringify(R) + 'FIN';
})();
