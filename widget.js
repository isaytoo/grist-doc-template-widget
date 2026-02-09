/**
 * Grist Document Template (Mail Merge) Widget
 * WYSIWYG editor with Grist field variables, mail merge preview,
 * Word import, and PDF generation.
 *
 * @author Said Hamadou
 * @license Apache-2.0
 * @copyright 2026 Said Hamadou
 */

// =============================================================================
// i18n
// =============================================================================

var currentLang = 'fr';

var i18n = {
  fr: {
    title: 'Document Template',
    notInGrist: 'Ce widget doit être utilisé dans Grist.',
    tabEditor: 'Éditeur',
    tabPreview: 'Prévisualisation',
    tabPdf: 'Générer PDF',
    selectTable: '📊 Table source :',
    selectTableOption: '-- Choisir une table --',
    importTitle: 'Importer un document Word',
    importDrop: 'Glissez un fichier .docx ici ou cliquez pour parcourir',
    importSuccess: 'Document Word importé avec succès !',
    importError: 'Erreur lors de l\'import : ',
    varTitle: 'Variables disponibles',
    varHint: '(cliquez pour insérer)',
    saveTemplate: 'Sauvegarder le modèle',
    clearEditor: 'Vider l\'éditeur',
    templateSaved: 'Modèle sauvegardé !',
    templateCleared: 'Éditeur vidé.',
    templateLoaded: 'Modèle chargé depuis la sauvegarde.',
    recordLabel: 'Enregistrement',
    previewEmpty: 'Sélectionnez une table et créez un modèle pour voir la prévisualisation.',
    pdfSingle: 'PDF de cet enregistrement',
    pdfTitle: 'Générer les PDF',
    pdfDesc: 'Générez un PDF pour chaque enregistrement de la table, ou un PDF combiné.',
    pdfFilename: 'Nom du fichier :',
    pdfMode: 'Mode :',
    pdfModeAll: 'Tous les enregistrements (1 PDF)',
    pdfModeCurrent: 'Enregistrement actuel uniquement',
    pdfPageSize: 'Format :',
    pdfGenerate: 'Générer le PDF',
    pdfGenerating: 'Génération en cours... {current}/{total}',
    pdfDone: 'PDF généré avec succès ! ({count} pages)',
    pdfError: 'Erreur lors de la génération : ',
    footerCreated: 'Créé par',
    confirmClear: 'Voulez-vous vraiment vider l\'éditeur ?',
    confirmClearTitle: 'Vider l\'éditeur',
    cancel: 'Annuler',
    confirm: 'Confirmer',
    pdfCancel: 'Annuler (conserver le partiel)',
    pdfCancelled: 'Génération annulée. PDF partiel sauvegardé ({count} pages).',
    noTemplate: 'Aucun modèle. Créez d\'abord un document dans l\'onglet Éditeur.',
    noData: 'Aucune donnée dans la table sélectionnée.',
    editorPlaceholder: '<p style="color:#94a3b8;">Commencez à écrire votre document ici... Utilisez les variables ci-dessus pour insérer des champs dynamiques.</p>'
  },
  en: {
    title: 'Document Template',
    notInGrist: 'This widget must be used inside Grist.',
    tabEditor: 'Editor',
    tabPreview: 'Preview',
    tabPdf: 'Generate PDF',
    selectTable: '📊 Source table:',
    selectTableOption: '-- Choose a table --',
    importTitle: 'Import a Word document',
    importDrop: 'Drag a .docx file here or click to browse',
    importSuccess: 'Word document imported successfully!',
    importError: 'Error importing: ',
    varTitle: 'Available variables',
    varHint: '(click to insert)',
    saveTemplate: 'Save template',
    clearEditor: 'Clear editor',
    templateSaved: 'Template saved!',
    templateCleared: 'Editor cleared.',
    templateLoaded: 'Template loaded from saved data.',
    recordLabel: 'Record',
    previewEmpty: 'Select a table and create a template to see the preview.',
    pdfSingle: 'PDF of this record',
    pdfTitle: 'Generate PDFs',
    pdfDesc: 'Generate a PDF for each record in the table, or a combined PDF.',
    pdfFilename: 'File name:',
    pdfMode: 'Mode:',
    pdfModeAll: 'All records (1 PDF)',
    pdfModeCurrent: 'Current record only',
    pdfPageSize: 'Page size:',
    pdfGenerate: 'Generate PDF',
    pdfGenerating: 'Generating... {current}/{total}',
    pdfDone: 'PDF generated successfully! ({count} pages)',
    pdfError: 'Error generating: ',
    footerCreated: 'Created by',
    confirmClear: 'Do you really want to clear the editor?',
    confirmClearTitle: 'Clear editor',
    cancel: 'Cancel',
    confirm: 'Confirm',
    pdfCancel: 'Cancel (keep partial)',
    pdfCancelled: 'Generation cancelled. Partial PDF saved ({count} pages).',
    noTemplate: 'No template. Create a document in the Editor tab first.',
    noData: 'No data in the selected table.',
    editorPlaceholder: '<p style="color:#94a3b8;">Start writing your document here... Use the variables above to insert dynamic fields.</p>'
  }
};

function t(key) {
  return (i18n[currentLang] && i18n[currentLang][key]) || (i18n.fr[key]) || key;
}

function setLang(lang) {
  currentLang = lang;
  document.querySelectorAll('.lang-btn').forEach(function(btn) {
    btn.classList.toggle('active', btn.textContent.trim() === lang.toUpperCase());
  });
  document.querySelectorAll('[data-i18n]').forEach(function(el) {
    var key = el.getAttribute('data-i18n');
    if (i18n[lang][key]) el.textContent = i18n[lang][key];
  });
}

// =============================================================================
// UTILS
// =============================================================================

function sanitize(str) {
  var div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

function isInsideGrist() {
  try { return window.self !== window.top; } catch (e) { return true; }
}

// =============================================================================
// TOAST & MODAL
// =============================================================================

function showToast(message, type) {
  var container = document.getElementById('toast-container');
  var toast = document.createElement('div');
  toast.className = 'toast toast-' + (type || 'info');
  toast.textContent = message;
  container.appendChild(toast);
  setTimeout(function() {
    toast.style.opacity = '0';
    toast.style.transition = 'opacity 0.3s';
    setTimeout(function() { toast.remove(); }, 300);
  }, 3500);
}

var modalResolve = null;
function showModal(title, body) {
  document.getElementById('modal-title').textContent = title;
  document.getElementById('modal-body').innerHTML = body;
  document.getElementById('modal-overlay').classList.remove('hidden');
  return new Promise(function(resolve) {
    modalResolve = resolve;
  });
}
function closeModal(result) {
  document.getElementById('modal-overlay').classList.add('hidden');
  if (modalResolve) { modalResolve(result || false); modalResolve = null; }
}

// =============================================================================
// STATE
// =============================================================================

var allTables = [];
var selectedTable = '';
var tableColumns = [];
var tableData = null;       // { col: [...], ... }
var currentRecordIndex = 0;
var templateHtml = '';
var editorInstance = null;
var TEMPLATE_STORAGE_KEY = 'grist_doc_template_';
var pdfCancelled = false;

// =============================================================================
// TABS
// =============================================================================

function switchTab(tabId) {
  document.querySelectorAll('.tab-btn').forEach(function(btn) {
    btn.classList.toggle('active', btn.getAttribute('data-tab') === tabId);
  });
  document.querySelectorAll('.tab-content').forEach(function(tc) {
    tc.classList.toggle('active', tc.id === 'tab-' + tabId);
  });
  if (tabId === 'preview') {
    renderPreview();
  }
}

// =============================================================================
// INIT
// =============================================================================

if (!isInsideGrist()) {
  document.getElementById('not-in-grist').classList.remove('hidden');
  document.getElementById('main-content').classList.add('hidden');
} else {
  (async function init() {
    try {
      await grist.ready({ requiredAccess: 'full' });
      console.log('Doc Template widget ready');
      await loadTables();
      initEditor();
    } catch (error) {
      console.error('Init error:', error);
    }
  })();
}

// =============================================================================
// LOAD TABLES
// =============================================================================

async function loadTables() {
  var loading = document.getElementById('table-loading');
  loading.classList.remove('hidden');
  try {
    var tables = await grist.docApi.listTables();
    allTables = tables.filter(function(t) {
      return !t.startsWith('_grist_') && !t.startsWith('GristHidden_');
    });
    var select = document.getElementById('table-select');
    select.innerHTML = '<option value="">' + t('selectTableOption') + '</option>';
    for (var i = 0; i < allTables.length; i++) {
      var opt = document.createElement('option');
      opt.value = allTables[i];
      opt.textContent = allTables[i];
      select.appendChild(opt);
    }
  } catch (error) {
    console.error('Error loading tables:', error);
  } finally {
    loading.classList.add('hidden');
  }
}

async function onTableChange() {
  var select = document.getElementById('table-select');
  selectedTable = select.value;
  if (!selectedTable) {
    tableColumns = [];
    tableData = null;
    document.getElementById('var-panel').classList.add('hidden');
    return;
  }

  try {
    var data = await grist.docApi.fetchTable(selectedTable);
    tableData = data;
    tableColumns = Object.keys(data).filter(function(c) {
      return c !== 'id' && c !== 'manualSort' && !c.startsWith('gristHelper_');
    });
    renderVariableChips();
    document.getElementById('var-panel').classList.remove('hidden');

    // Load saved template for this table
    loadSavedTemplate();

    currentRecordIndex = 0;
  } catch (error) {
    console.error('Error loading table data:', error);
    showToast(t('importError') + error.message, 'error');
  }
}

// =============================================================================
// VARIABLE CHIPS
// =============================================================================

function renderVariableChips() {
  var html = '';
  for (var i = 0; i < tableColumns.length; i++) {
    var col = tableColumns[i];
    html += '<span class="var-chip" onclick="insertVariable(\'' + sanitize(col) + '\')">';
    html += '{{' + sanitize(col) + '}}';
    html += '</span>';
  }
  document.getElementById('var-chips').innerHTML = html;
}

function insertVariable(colName) {
  if (!editorInstance) return;
  var varHtml = '<span style="background:#f3e8ff;color:#7c3aed;padding:2px 6px;border-radius:4px;font-weight:600;" contenteditable="false">{{' + colName + '}}</span>&nbsp;';
  editorInstance.selection.insertHTML(varHtml);
  showToast('{{' + colName + '}} inséré', 'info');
}

// =============================================================================
// JODIT EDITOR
// =============================================================================

function initEditor() {
  editorInstance = Jodit.make('#editor-container', {
    language: currentLang,
    height: 500,
    placeholder: currentLang === 'fr' ? 'Commencez à écrire votre document ici...' : 'Start writing your document here...',
    toolbarAdaptive: false,
    askBeforePasteHTML: false,
    askBeforePasteFromWord: false,
    defaultActionOnPaste: 'insert_clear_html',
    buttons: [
      'bold', 'italic', 'underline', 'strikethrough', '|',
      'font', 'fontsize', '|',
      'brush', '|',
      'paragraph', '|',
      'ul', 'ol', '|',
      'outdent', 'indent', '|',
      'align', '|',
      'table', '|',
      'link', 'image', '|',
      'hr', 'symbol', '|',
      'undo', 'redo', '|',
      'eraser', 'source', 'fullsize', 'print'
    ],
    style: {
      'font-family': '"Times New Roman", Times, serif',
      'font-size': '14px',
      'line-height': '1.6'
    },
    iframe: false,
    showCharsCounter: false,
    showWordsCounter: false,
    showXPathInStatusbar: false
  });
}

// =============================================================================
// SAVE / LOAD TEMPLATE
// =============================================================================

function getEditorHtml() {
  if (!editorInstance) return '';
  return editorInstance.value;
}

function setEditorHtml(html) {
  if (!editorInstance) return;
  editorInstance.value = html;
}

function saveTemplate() {
  if (!editorInstance) return;
  templateHtml = getEditorHtml();
  if (selectedTable) {
    try {
      localStorage.setItem(TEMPLATE_STORAGE_KEY + selectedTable, templateHtml);
    } catch (e) { /* localStorage may not be available in iframe */ }
  }
  showToast(t('templateSaved'), 'success');
}

function loadSavedTemplate() {
  try {
    var saved = localStorage.getItem(TEMPLATE_STORAGE_KEY + selectedTable);
    if (saved && editorInstance) {
      setEditorHtml(saved);
      templateHtml = saved;
      showToast(t('templateLoaded'), 'info');
    }
  } catch (e) { /* localStorage may not be available */ }
}

async function clearEditor() {
  var confirmed = await showModal(t('confirmClearTitle'), t('confirmClear'));
  if (confirmed && editorInstance) {
    editorInstance.value = '';
    templateHtml = '';
    if (selectedTable) {
      try { localStorage.removeItem(TEMPLATE_STORAGE_KEY + selectedTable); } catch (e) {}
    }
    showToast(t('templateCleared'), 'info');
  }
}

// =============================================================================
// IMPORT WORD (.docx)
// =============================================================================

// Drag & drop
var dropZone = document.getElementById('drop-zone');
if (dropZone) {
  dropZone.addEventListener('dragover', function(e) {
    e.preventDefault();
    dropZone.classList.add('dragover');
  });
  dropZone.addEventListener('dragleave', function() {
    dropZone.classList.remove('dragover');
  });
  dropZone.addEventListener('drop', function(e) {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    var files = e.dataTransfer.files;
    if (files.length > 0 && files[0].name.endsWith('.docx')) {
      importWord(files[0]);
    }
  });
}

function importWord(file) {
  if (!file) return;
  showToast(currentLang === 'fr' ? 'Import en cours...' : 'Importing...', 'info');
  var reader = new FileReader();
  reader.onload = function(e) {
    var arrayBuffer = e.target.result;
    var options = {
      styleMap: [
        "p[style-name='Title'] => h1:fresh",
        "p[style-name='Titre'] => h1:fresh",
        "p[style-name='Heading 1'] => h1:fresh",
        "p[style-name='Titre 1'] => h1:fresh",
        "p[style-name='Heading 2'] => h2:fresh",
        "p[style-name='Titre 2'] => h2:fresh",
        "p[style-name='Heading 3'] => h3:fresh",
        "p[style-name='Titre 3'] => h3:fresh",
        "p[style-name='Heading 4'] => h4:fresh",
        "p[style-name='Titre 4'] => h4:fresh",
        "p[style-name='List Paragraph'] => li:fresh",
        "p[style-name='Paragraphe de liste'] => li:fresh",
        "p[style-name='Quote'] => blockquote:fresh",
        "p[style-name='Citation'] => blockquote:fresh",
        "p[style-name='Intense Quote'] => blockquote:fresh",
        "p[style-name='Subtitle'] => h2:fresh",
        "p[style-name='Sous-titre'] => h2:fresh",
        "r[style-name='Strong'] => strong",
        "r[style-name='Emphasis'] => em",
        "b => strong",
        "i => em",
        "u => u",
        "strike => s"
      ],
      convertImage: mammoth.images.imgElement(function(image) {
        return image.read('base64').then(function(imageBuffer) {
          return {
            src: 'data:' + image.contentType + ';base64,' + imageBuffer
          };
        });
      }),
      includeDefaultStyleMap: true
    };
    mammoth.convertToHtml({ arrayBuffer: arrayBuffer }, options).then(function(result) {
      var html = result.value;
      if (result.messages.length > 0) {
        console.log('Mammoth messages:', result.messages);
      }
      // Post-process HTML for better layout
      html = postProcessWordHtml(html);
      if (editorInstance) {
        setEditorHtml(html);
        templateHtml = html;
      }
      var warnings = result.messages.filter(function(m) { return m.type === 'warning'; });
      if (warnings.length > 0) {
        showToast(t('importSuccess') + ' (' + warnings.length + ' avertissements)', 'warning');
      } else {
        showToast(t('importSuccess'), 'success');
      }
    }).catch(function(error) {
      console.error('Word import error:', error);
      showToast(t('importError') + error.message, 'error');
    });
  };
  reader.readAsArrayBuffer(file);
}

function postProcessWordHtml(html) {
  // Add max-width to images so they fit in the editor
  html = html.replace(/<img /g, '<img style="max-width:100%;height:auto;" ');

  // Convert page breaks (Mammoth doesn't handle them well)
  // Some Word docs use <br/> for page breaks - we add a visual separator
  html = html.replace(/<br\s*\/?>\s*<br\s*\/?>\s*<br\s*\/?>/g,
    '<hr style="border:none;border-top:2px dashed #cbd5e1;margin:30px 0;page-break-after:always;">');

  // Ensure empty paragraphs have some height (Word uses them for spacing)
  html = html.replace(/<p><\/p>/g, '<p style="min-height:1em;">&nbsp;</p>');

  // Add some spacing to paragraphs
  html = html.replace(/<p>/g, '<p style="margin-bottom:8px;">');

  // Style tables - no visible borders, just padding for layout
  html = html.replace(/<table>/g, '<table style="border-collapse:collapse;width:100%;margin:10px 0;border:none;">');
  html = html.replace(/<td>/g, '<td style="border:none;padding:6px 10px;vertical-align:top;">');
  html = html.replace(/<td /g, '<td style="border:none;padding:6px 10px;vertical-align:top;" ');
  html = html.replace(/<th>/g, '<th style="border:none;padding:6px 10px;font-weight:bold;vertical-align:top;">');
  html = html.replace(/<th /g, '<th style="border:none;padding:6px 10px;font-weight:bold;vertical-align:top;" ');

  return html;
}

// =============================================================================
// PREVIEW (MAIL MERGE)
// =============================================================================

function getRecordCount() {
  if (!tableData || !tableColumns.length) return 0;
  var firstCol = tableColumns[0];
  return (tableData[firstCol] || []).length;
}

function getRecordAt(index) {
  if (!tableData || index < 0) return {};
  var record = {};
  for (var i = 0; i < tableColumns.length; i++) {
    var col = tableColumns[i];
    var arr = tableData[col] || [];
    record[col] = (index < arr.length) ? arr[index] : '';
  }
  return record;
}

function resolveTemplate(html, record, forPdf) {
  var resolved = html;
  for (var col in record) {
    var val = record[col];
    var display = (val === null || val === undefined || val === '') ? '' : String(val);
    // Replace styled spans (Quill wraps variables in <span style="...">)
    var styledRegex = new RegExp('<span[^>]*>\\{\\{' + escapeRegex(col) + '\\}\\}</span>', 'g');
    if (display) {
      if (forPdf) {
        resolved = resolved.replace(styledRegex, '<strong>' + sanitize(display) + '</strong>');
      } else {
        resolved = resolved.replace(styledRegex, '<span class="var-resolved">' + sanitize(display) + '</span>');
      }
    } else {
      if (forPdf) {
        resolved = resolved.replace(styledRegex, '<em>[' + col + ']</em>');
      } else {
        resolved = resolved.replace(styledRegex, '<span class="var-empty">[' + col + ': vide]</span>');
      }
    }
    // Replace plain text {{col}}
    var plainRegex = new RegExp('\\{\\{' + escapeRegex(col) + '\\}\\}', 'g');
    if (display) {
      if (forPdf) {
        resolved = resolved.replace(plainRegex, '<strong>' + sanitize(display) + '</strong>');
      } else {
        resolved = resolved.replace(plainRegex, '<span class="var-resolved">' + sanitize(display) + '</span>');
      }
    } else {
      if (forPdf) {
        resolved = resolved.replace(plainRegex, '<em>[' + col + ']</em>');
      } else {
        resolved = resolved.replace(plainRegex, '<span class="var-empty">[' + col + ': vide]</span>');
      }
    }
  }
  // For PDF: strip variable styling and table borders
  if (forPdf) {
    resolved = resolved.replace(/background-color:\s*rgb\(243,\s*232,\s*255\);?/g, '');
    resolved = resolved.replace(/background-color:\s*#f3e8ff;?/g, '');
    resolved = resolved.replace(/color:\s*rgb\(124,\s*58,\s*237\);?/g, '');
    resolved = resolved.replace(/color:\s*#7c3aed;?/g, '');
    // Strip all table/cell borders for clean PDF
    resolved = resolved.replace(/(<table[^>]*?)style="[^"]*"/g, function(m, pre) {
      return pre + 'style="border-collapse:collapse;width:100%;border:none;"';
    });
    resolved = resolved.replace(/(<td[^>]*?)style="[^"]*"/g, function(m, pre) {
      return pre + 'style="border:none;padding:6px 10px;vertical-align:top;"';
    });
    resolved = resolved.replace(/(<th[^>]*?)style="[^"]*"/g, function(m, pre) {
      return pre + 'style="border:none;padding:6px 10px;font-weight:bold;vertical-align:top;"';
    });
    // Also handle tables/cells without style attribute
    resolved = resolved.replace(/<table(?![^>]*style=)/g, '<table style="border-collapse:collapse;width:100%;border:none;"');
    resolved = resolved.replace(/<td(?![^>]*style=)/g, '<td style="border:none;padding:6px 10px;vertical-align:top;"');
    resolved = resolved.replace(/<th(?![^>]*style=)/g, '<th style="border:none;padding:6px 10px;font-weight:bold;vertical-align:top;"');
  }
  return resolved;
}

function escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function renderPreview() {
  // Get current template from editor
  if (editorInstance) {
    templateHtml = getEditorHtml();
  }

  var count = getRecordCount();
  document.getElementById('record-total').textContent = count;

  if (!templateHtml || !selectedTable || count === 0) {
    document.getElementById('preview-page').innerHTML =
      '<p style="color:#94a3b8; text-align:center; padding:40px;">' + t('previewEmpty') + '</p>';
    document.getElementById('record-current').textContent = '0';
    return;
  }

  if (currentRecordIndex >= count) currentRecordIndex = count - 1;
  if (currentRecordIndex < 0) currentRecordIndex = 0;

  document.getElementById('record-current').textContent = currentRecordIndex + 1;

  var record = getRecordAt(currentRecordIndex);
  var resolved = resolveTemplate(templateHtml, record);
  document.getElementById('preview-page').innerHTML = resolved;
}

function prevRecord() {
  if (currentRecordIndex > 0) {
    currentRecordIndex--;
    renderPreview();
  }
}

function nextRecord() {
  var count = getRecordCount();
  if (currentRecordIndex < count - 1) {
    currentRecordIndex++;
    renderPreview();
  }
}

// =============================================================================
// PDF GENERATION
// =============================================================================

async function generateSinglePdf() {
  if (!templateHtml || !selectedTable) {
    showToast(t('noTemplate'), 'error');
    return;
  }
  var count = getRecordCount();
  if (count === 0) {
    showToast(t('noData'), 'error');
    return;
  }

  var record = getRecordAt(currentRecordIndex);
  var resolved = resolveTemplate(templateHtml, record, true);
  await generatePdfFromHtml(resolved, 'document_' + (currentRecordIndex + 1) + '.pdf');
}

function cancelPdf() {
  pdfCancelled = true;
}

async function generatePdf() {
  if (!templateHtml && editorInstance) {
    templateHtml = getEditorHtml();
  }
  if (!templateHtml || !selectedTable) {
    showToast(t('noTemplate'), 'error');
    return;
  }
  var count = getRecordCount();
  if (count === 0) {
    showToast(t('noData'), 'error');
    return;
  }

  var mode = document.getElementById('pdf-mode').value;
  var filename = document.getElementById('pdf-filename').value.trim() || 'document_publipostage';
  var pageSize = document.getElementById('pdf-page-size').value;

  var btn = document.getElementById('pdf-generate-btn');
  var cancelBtn = document.getElementById('pdf-cancel-btn');
  btn.disabled = true;
  cancelBtn.classList.remove('hidden');
  pdfCancelled = false;
  var progressBar = document.getElementById('pdf-progress');
  var progressFill = document.getElementById('pdf-progress-fill');
  var messageDiv = document.getElementById('pdf-message');
  progressBar.classList.remove('hidden');
  messageDiv.classList.remove('hidden');

  var pagesGenerated = 0;

  try {
    var jsPDF = window.jspdf.jsPDF;
    var orientation = 'portrait';
    var format = pageSize === 'a4' ? 'a4' : 'letter';

    var startIdx = 0;
    var endIdx = count;
    if (mode === 'current') {
      startIdx = currentRecordIndex;
      endIdx = currentRecordIndex + 1;
    }

    var totalPages = endIdx - startIdx;
    var pdf = new jsPDF({ orientation: orientation, unit: 'mm', format: format });
    var pageWidth = pdf.internal.pageSize.getWidth();
    var pageHeight = pdf.internal.pageSize.getHeight();

    for (var i = startIdx; i < endIdx; i++) {
      // Check cancel
      if (pdfCancelled) break;

      var progress = Math.round(((i - startIdx + 1) / totalPages) * 100);
      progressFill.style.width = progress + '%';
      messageDiv.innerHTML = '<div class="message message-info">' +
        t('pdfGenerating').replace('{current}', i - startIdx + 1).replace('{total}', totalPages) + '</div>';

      var record = getRecordAt(i);
      var resolved = resolveTemplate(templateHtml, record, true);

      // Render to a temporary div
      var tempDiv = document.createElement('div');
      tempDiv.style.cssText = 'position:absolute;left:-9999px;top:0;width:' + (pageSize === 'a4' ? '794' : '816') + 'px;padding:40px 60px;font-family:"Times New Roman",Times,serif;font-size:14px;line-height:1.6;background:white;';
      tempDiv.innerHTML = resolved;
      document.body.appendChild(tempDiv);

      // Wait for layout to settle
      await new Promise(function(resolve) { setTimeout(resolve, 100); });

      // Check cancel again after wait
      if (pdfCancelled) {
        document.body.removeChild(tempDiv);
        break;
      }

      // Use html2canvas
      var canvas = await html2canvas(tempDiv, {
        scale: 2,
        useCORS: true,
        logging: false,
        backgroundColor: '#ffffff',
        scrollX: 0,
        scrollY: 0,
        windowWidth: tempDiv.scrollWidth,
        windowHeight: tempDiv.scrollHeight
      });

      document.body.removeChild(tempDiv);

      var imgData = canvas.toDataURL('image/jpeg', 0.95);
      var imgWidth = pageWidth - 20; // 10mm margin each side
      var imgHeight = (canvas.height * imgWidth) / canvas.width;

      // Handle multi-page content
      var yOffset = 0;
      var availableHeight = pageHeight - 20; // 10mm margin top/bottom

      if (i > startIdx) {
        pdf.addPage();
      }

      while (yOffset < imgHeight) {
        if (yOffset > 0) {
          pdf.addPage();
        }
        // Calculate source crop
        var sourceY = (yOffset / imgHeight) * canvas.height;
        var sourceH = Math.min((availableHeight / imgHeight) * canvas.height, canvas.height - sourceY);
        var drawH = (sourceH / canvas.height) * imgHeight;

        // Create cropped canvas
        var cropCanvas = document.createElement('canvas');
        cropCanvas.width = canvas.width;
        cropCanvas.height = sourceH;
        var ctx = cropCanvas.getContext('2d');
        ctx.drawImage(canvas, 0, sourceY, canvas.width, sourceH, 0, 0, canvas.width, sourceH);

        var cropImgData = cropCanvas.toDataURL('image/jpeg', 0.95);
        pdf.addImage(cropImgData, 'JPEG', 10, 10, imgWidth, drawH);

        yOffset += availableHeight;
      }

      pagesGenerated = i - startIdx + 1;

      // Yield to UI
      await new Promise(function(resolve) { setTimeout(resolve, 50); });
    }

    // Save (full or partial)
    pdf.save(filename + '.pdf');

    if (pdfCancelled && pagesGenerated > 0) {
      progressFill.style.width = Math.round((pagesGenerated / totalPages) * 100) + '%';
      messageDiv.innerHTML = '<div class="message message-warning" style="background:#fffbeb;color:#92400e;border:1px solid #fde68a;">' +
        t('pdfCancelled').replace('{count}', pagesGenerated) + '</div>';
      showToast(t('pdfCancelled').replace('{count}', pagesGenerated), 'warning');
    } else {
      progressFill.style.width = '100%';
      messageDiv.innerHTML = '<div class="message message-success">' +
        t('pdfDone').replace('{count}', totalPages) + '</div>';
      showToast(t('pdfDone').replace('{count}', totalPages), 'success');
    }

  } catch (error) {
    console.error('PDF generation error:', error);
    messageDiv.innerHTML = '<div class="message message-error">' + t('pdfError') + sanitize(error.message) + '</div>';
    showToast(t('pdfError') + error.message, 'error');
  } finally {
    btn.disabled = false;
    cancelBtn.classList.add('hidden');
    pdfCancelled = false;
  }
}

async function generatePdfFromHtml(html, filename) {
  try {
    var jsPDF = window.jspdf.jsPDF;
    var pageSize = document.getElementById('pdf-page-size').value;
    var format = pageSize === 'a4' ? 'a4' : 'letter';

    var pdf = new jsPDF({ orientation: 'portrait', unit: 'mm', format: format });
    var pageWidth = pdf.internal.pageSize.getWidth();
    var pageHeight = pdf.internal.pageSize.getHeight();

    var tempDiv = document.createElement('div');
    tempDiv.style.cssText = 'position:absolute;left:-9999px;top:0;width:' + (pageSize === 'a4' ? '794' : '816') + 'px;padding:40px 60px;font-family:"Times New Roman",Times,serif;font-size:14px;line-height:1.6;background:white;';
    tempDiv.innerHTML = html;
    document.body.appendChild(tempDiv);

    // Wait for layout to settle
    await new Promise(function(resolve) { setTimeout(resolve, 100); });

    var canvas = await html2canvas(tempDiv, {
      scale: 2,
      useCORS: true,
      logging: false,
      backgroundColor: '#ffffff',
      scrollX: 0,
      scrollY: 0,
      windowWidth: tempDiv.scrollWidth,
      windowHeight: tempDiv.scrollHeight
    });

    document.body.removeChild(tempDiv);

    var imgWidth = pageWidth - 20;
    var imgHeight = (canvas.height * imgWidth) / canvas.width;
    var availableHeight = pageHeight - 20;
    var yOffset = 0;

    while (yOffset < imgHeight) {
      if (yOffset > 0) pdf.addPage();
      var sourceY = (yOffset / imgHeight) * canvas.height;
      var sourceH = Math.min((availableHeight / imgHeight) * canvas.height, canvas.height - sourceY);
      var drawH = (sourceH / canvas.height) * imgHeight;

      var cropCanvas = document.createElement('canvas');
      cropCanvas.width = canvas.width;
      cropCanvas.height = sourceH;
      var ctx = cropCanvas.getContext('2d');
      ctx.drawImage(canvas, 0, sourceY, canvas.width, sourceH, 0, 0, canvas.width, sourceH);

      var cropImgData = cropCanvas.toDataURL('image/jpeg', 0.95);
      pdf.addImage(cropImgData, 'JPEG', 10, 10, imgWidth, drawH);

      yOffset += availableHeight;
    }

    pdf.save(filename);
    showToast('PDF généré !', 'success');

  } catch (error) {
    console.error('PDF error:', error);
    showToast(t('pdfError') + error.message, 'error');
  }
}
