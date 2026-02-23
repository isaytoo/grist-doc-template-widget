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
var columnMetadata = {}; // Store column metadata (type, reference info)
var referenceTables = {}; // Cache for referenced table data
var referenceDisplayValues = {}; // Map ref table -> id -> display value (e.g., PARAMETRES -> 60 -> "DUMZ 60")
var columnIdToName = {}; // Map column ID -> column name (for resolving visibleCol)

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
    editorPlaceholder: '<p style="color:#94a3b8;">Commencez à écrire votre document ici... Utilisez les variables ci-dessus pour insérer des champs dynamiques.</p>',
    templateName: 'Nom du modèle :',
    templateNamePlaceholder: 'Ex: Courrier standard, PV réunion...',
    templateSelect: 'Modèles enregistrés :',
    templateSelectDefault: '-- Nouveau modèle --',
    templateDeleteConfirm: 'Supprimer le modèle "{name}" ?',
    templateDeleted: 'Modèle "{name}" supprimé.',
    templateDelete: 'Supprimer',
    templateLoad: 'Charger',
    loopHint: 'Boucle (plusieurs lignes)',
    loopSyntax: '{{#each Colonne=Valeur}}...{{/each}}',
    loopExample: 'Ex: {{#each Date=16/02/26}}{{Prenom}}<br>{{/each}}',
    tableLoopBtn: 'Tableau avec boucle',
    tableLoopHint: 'Insérer un tableau qui répète les lignes filtrées',
    tableLoopInserted: 'Tableau avec boucle inséré',
    importBtn: 'Importer .docx',
    boolFormat: 'Cases à cocher :',
    boolFormatText: 'Oui / Non',
    boolFormatCheckbox: '☑ / ☐',
    skipEmptyPages: 'Pages vides :',
    skipEmptyPagesYes: 'Ignorer (recto seul)',
    skipEmptyPagesNo: 'Conserver (recto-verso)'
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
    boolFormat: 'Checkboxes:',
    boolFormatText: 'Yes / No',
    boolFormatCheckbox: '☑ / ☐',
    skipEmptyPages: 'Empty pages:',
    skipEmptyPagesYes: 'Skip (single-sided)',
    skipEmptyPagesNo: 'Keep (double-sided)',
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
    editorPlaceholder: '<p style="color:#94a3b8;">Start writing your document here... Use the variables above to insert dynamic fields.</p>',
    templateName: 'Template name:',
    templateNamePlaceholder: 'E.g.: Standard letter, Meeting notes...',
    templateSelect: 'Saved templates:',
    templateSelectDefault: '-- New template --',
    templateDeleteConfirm: 'Delete template "{name}"?',
    templateDeleted: 'Template "{name}" deleted.',
    templateDelete: 'Delete',
    templateLoad: 'Load',
    loopHint: 'Loop (multiple rows)',
    loopSyntax: '{{#each Column=Value}}...{{/each}}',
    loopExample: 'Ex: {{#each Date=16/02/26}}{{FirstName}}<br>{{/each}}',
    tableLoopBtn: 'Table with loop',
    tableLoopHint: 'Insert a table that repeats filtered rows',
    tableLoopInserted: 'Table with loop inserted',
    importBtn: 'Import .docx'
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
var tableData = null;       // { col: [...], ... } - ALL data from table
var filteredRecords = [];   // Records filtered by "Select By" (from grist.onRecords)
var availableViews = [];    // Views available for the selected table
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
  
  // Show/hide fixed bars based on tab
  var fixedVarBar = document.getElementById('fixed-var-bar');
  var fixedBottomBar = document.querySelector('.fixed-bottom-bar');
  if (tabId === 'editor') {
    if (fixedVarBar) fixedVarBar.style.display = 'block';
    if (fixedBottomBar) fixedBottomBar.style.display = 'block';
  } else {
    if (fixedVarBar) fixedVarBar.style.display = 'none';
    if (fixedBottomBar) fixedBottomBar.style.display = 'none';
  }
  
  if (tabId === 'preview') {
    renderPreview();
  }
  if (tabId === 'pdf') {
    refreshPdfTemplateList();
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

      // Initialize editor FIRST so it's ready when we load templates
      initEditor();
      
      // Show fixed bars for editor tab (default tab)
      var fixedVarBar = document.getElementById('fixed-var-bar');
      var fixedBottomBar = document.querySelector('.fixed-bottom-bar');
      if (fixedVarBar) fixedVarBar.style.display = 'block';
      if (fixedBottomBar) fixedBottomBar.style.display = 'block';

      // Restore draft immediately after editor init
      try {
        var draft = await grist.widgetApi.getOption('editorDraft');
        console.log('Draft from options:', draft ? draft.substring(0, 50) + '...' : 'null');
        if (draft && editorInstance) {
          setEditorHtml(draft);
          templateHtml = draft;
          console.log('Draft restored at startup');
        }
      } catch (e) {
        console.warn('Could not restore draft:', e);
      }

      // Listen for widget options (template stored in Grist)
      grist.onOptions(function(options) {
        if (options && options.template && selectedTable) {
          var key = 'template_' + selectedTable;
          if (options[key]) {
            setEditorHtml(options[key]);
            templateHtml = options[key];
            console.log('Template loaded from Grist options for', selectedTable);
          }
        }
      });

      // Listen for filtered records from "Select By" configuration
      grist.onRecords(function(records) {
        console.log('Received filtered records from Select By:', records.length);
        filteredRecords = records || [];
      });

      // Load tables and restore selection
      await loadTables();
      
      // Load saved templates list
      await refreshTemplateList();
      console.log('Template list refreshed at startup');
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
    
    // Restore previously selected table
    try {
      var savedTable = await grist.widgetApi.getOption('selectedTable');
      if (savedTable && allTables.indexOf(savedTable) !== -1) {
        select.value = savedTable;
        await onTableChange(true); // true = skip saving again
      }
    } catch (e) {
      console.warn('Could not restore selected table:', e);
    }
  } catch (error) {
    console.error('Error loading tables:', error);
  } finally {
    loading.classList.add('hidden');
  }
}

async function onTableChange(skipSave) {
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
    
    // Fetch column metadata to detect Reference columns
    await loadColumnMetadata(selectedTable);
    // Resolve reference values
    await resolveReferences();
    // Load available views for this table
    await loadViewsForTable(selectedTable);
    
    renderVariableChips();
    document.getElementById('var-panel').classList.remove('hidden');

    // Save selected table to widget options (persist across page changes)
    if (!skipSave) {
      try {
        await grist.widgetApi.setOption('selectedTable', selectedTable);
      } catch (e) {
        console.warn('Could not save selected table:', e);
      }
    }

    // Load saved template for this table (only if editor is empty)
    if (!getEditorHtml().trim()) {
      loadSavedTemplate();
    }

    currentRecordIndex = 0;
  } catch (error) {
    console.error('Error loading table data:', error);
    showToast(t('importError') + error.message, 'error');
  }
}

// Auto-save editor content periodically
var autoSaveTimer = null;
function scheduleAutoSave() {
  if (autoSaveTimer) clearTimeout(autoSaveTimer);
  autoSaveTimer = setTimeout(async function() {
    if (editorInstance) {
      var content = getEditorHtml();
      if (content && content.trim()) {
        try {
          await grist.widgetApi.setOption('editorDraft', content);
          console.log('Editor draft auto-saved');
        } catch (e) {}
      }
    }
  }, 2000); // Save 2 seconds after last change
}

// =============================================================================
// LOAD VIEWS FOR TABLE
// =============================================================================

async function loadViewsForTable(tableName) {
  availableViews = [];
  try {
    // Fetch views and view sections from Grist system tables
    var viewsData = await grist.docApi.fetchTable('_grist_Views');
    var sectionsData = await grist.docApi.fetchTable('_grist_Views_section');
    var tablesData = await grist.docApi.fetchTable('_grist_Tables');
    
    // Fetch _grist_Filters table (where Grist stores saved filters)
    var filtersData = null;
    try {
      filtersData = await grist.docApi.fetchTable('_grist_Filters');
    } catch (e) {
      console.log('_grist_Filters not available:', e.message);
    }
    
    // Find table ID for the selected table
    var tableId = null;
    if (tablesData && tablesData.id) {
      for (var i = 0; i < tablesData.id.length; i++) {
        if (tablesData.tableId[i] === tableName) {
          tableId = tablesData.id[i];
          break;
        }
      }
    }
    if (!tableId) return;
    
    // Build a map of sectionId -> filters from _grist_Filters
    var sectionFiltersFromGrist = {}; // sectionId -> [{colRef, filter}]
    if (filtersData && filtersData.id) {
      for (var i = 0; i < filtersData.id.length; i++) {
        var sectionRef = filtersData.viewSectionRef[i];
        var colRef = filtersData.colRef[i];
        var filterJson = filtersData.filter[i];
        
        if (!sectionFiltersFromGrist[sectionRef]) {
          sectionFiltersFromGrist[sectionRef] = [];
        }
        sectionFiltersFromGrist[sectionRef].push({
          colRef: colRef,
          filter: filterJson
        });
      }
    }
    
    // Find all view sections that use this table
    var viewIdsWithTable = new Set();
    var sectionFilters = {}; // sectionId -> filters array
    var sectionToView = {}; // sectionId -> viewId
    
    if (sectionsData && sectionsData.id) {
      for (var i = 0; i < sectionsData.id.length; i++) {
        if (sectionsData.tableRef && sectionsData.tableRef[i] === tableId) {
          var viewId = sectionsData.parentId ? sectionsData.parentId[i] : null;
          var sectionId = sectionsData.id[i];
          if (viewId) {
            viewIdsWithTable.add(viewId);
            sectionToView[sectionId] = viewId;
            
            // Get filters from _grist_Filters for this section
            if (sectionFiltersFromGrist[sectionId]) {
              sectionFilters[sectionId] = sectionFiltersFromGrist[sectionId];
            }
          }
        }
      }
    }
    
    // Build list of views with their names and filters
    if (viewsData && viewsData.id) {
      for (var i = 0; i < viewsData.id.length; i++) {
        var viewId = viewsData.id[i];
        if (viewIdsWithTable.has(viewId)) {
          var viewName = viewsData.name ? viewsData.name[i] : 'View ' + viewId;
          
          // Find filters for this view (from any of its sections)
          var viewFilters = null;
          var viewSectionId = null;
          for (var secId in sectionToView) {
            if (sectionToView[secId] === viewId && sectionFilters[secId]) {
              viewFilters = sectionFilters[secId];
              viewSectionId = secId;
              break;
            }
          }
          
          availableViews.push({
            id: viewId,
            name: viewName,
            filters: viewFilters,
            sectionId: viewSectionId
          });
        }
      }
    }
    
    console.log('Available views for table', tableName, ':', availableViews);
  } catch (error) {
    console.error('Error loading views:', error);
  }
}

// =============================================================================
// REFERENCE RESOLUTION
// =============================================================================

async function loadColumnMetadata(tableName) {
  columnMetadata = {};
  columnIdToName = {};
  try {
    // Fetch _grist_Tables_column to get column types
    var colData = await grist.docApi.fetchTable('_grist_Tables_column');
    var tablesData = await grist.docApi.fetchTable('_grist_Tables');
    
    // Build global column ID -> name mapping (for all tables)
    for (var j = 0; j < colData.id.length; j++) {
      columnIdToName[colData.id[j]] = colData.colId[j];
    }
    
    // Find table ID
    var tableId = null;
    for (var i = 0; i < tablesData.id.length; i++) {
      if (tablesData.tableId[i] === tableName) {
        tableId = tablesData.id[i];
        break;
      }
    }
    if (!tableId) return;
    
    // Get columns for this table
    for (var i = 0; i < colData.id.length; i++) {
      if (colData.parentId[i] === tableId) {
        var colId = colData.colId[i];
        var colType = colData.type[i];
        var displayCol = colData.displayCol[i];
        var visibleColId = colData.visibleCol[i];
        
        // Resolve visibleCol ID to column name
        var visibleColName = visibleColId ? columnIdToName[visibleColId] : null;
        
        columnMetadata[colId] = {
          type: colType,
          displayCol: displayCol,
          visibleCol: visibleColName // Now it's the column NAME, not ID
        };
      }
    }
    console.log('Column metadata loaded:', columnMetadata);
  } catch (e) {
    console.warn('Could not load column metadata:', e);
  }
}

async function resolveReferences() {
  if (!tableData || !columnMetadata) return;
  
  for (var colName in columnMetadata) {
    var meta = columnMetadata[colName];
    if (!meta.type) continue;
    
    // Check if it's a Reference or ReferenceList column
    var refMatch = meta.type.match(/^Ref:(.+)$/);
    var refListMatch = meta.type.match(/^RefList:(.+)$/);
    
    if (refMatch || refListMatch) {
      var refTableName = refMatch ? refMatch[1] : refListMatch[1];
      
      // Fetch the referenced table if not cached
      if (!referenceTables[refTableName]) {
        try {
          referenceTables[refTableName] = await grist.docApi.fetchTable(refTableName);
          console.log('Fetched reference table:', refTableName);
          
          // Build display values map using visibleCol from the reference column metadata
          // This is the column Grist uses to display reference values
          referenceDisplayValues[refTableName] = { byVisibleCol: {}, byFirstTextCol: {} };
          var refData = referenceTables[refTableName];
          
          // Use visibleCol from metadata if available, otherwise find display column
          var visibleColName = meta.visibleCol || findDisplayColumn(refData, null);
          
          // Also find the first text column (often contains identifiers like "DUMZ 60")
          var firstTextCol = null;
          for (var colKey in refData) {
            if (colKey !== 'id' && colKey !== 'manualSort' && !colKey.startsWith('gristHelper_')) {
              if (refData[colKey] && refData[colKey].length > 0 && typeof refData[colKey][0] === 'string') {
                firstTextCol = colKey;
                break;
              }
            }
          }
          
          if (refData.id) {
            for (var k = 0; k < refData.id.length; k++) {
              if (visibleColName && refData[visibleColName]) {
                referenceDisplayValues[refTableName].byVisibleCol[refData.id[k]] = refData[visibleColName][k];
              }
              if (firstTextCol && refData[firstTextCol] && firstTextCol !== visibleColName) {
                referenceDisplayValues[refTableName].byFirstTextCol[refData.id[k]] = refData[firstTextCol][k];
              }
            }
            console.log('Built reference display map for', refTableName, '- visibleCol:', visibleColName, ', firstTextCol:', firstTextCol);
          }
        } catch (e) {
          console.warn('Could not fetch reference table:', refTableName, e);
          continue;
        }
      }
      
      var refTable = referenceTables[refTableName];
      if (!refTable || !tableData[colName]) continue;
      
      // Find the display column (usually the first text column or rowId)
      var displayColName = findDisplayColumn(refTable, meta.visibleCol);
      
      // Replace IDs with display values
      var resolvedValues = [];
      for (var i = 0; i < tableData[colName].length; i++) {
        var refId = tableData[colName][i];
        if (refListMatch && Array.isArray(refId)) {
          // ReferenceList: array of IDs
          var names = [];
          for (var j = 0; j < refId.length; j++) {
            var name = lookupRefValue(refTable, refId[j], displayColName);
            if (name) names.push(name);
          }
          resolvedValues.push(names.join(', '));
        } else if (refId && typeof refId === 'number' && refId !== 0) {
          // Single Reference (0 means empty reference)
          var displayValue = lookupRefValue(refTable, refId, displayColName);
          resolvedValues.push(displayValue || refId);
        } else if (refId === 0 || refId === null || refId === undefined) {
          // Empty reference
          resolvedValues.push('');
        } else {
          resolvedValues.push(refId);
        }
      }
      tableData[colName] = resolvedValues;
      console.log('Resolved references for', colName, ':', resolvedValues.slice(0, 3));
    }
  }
}

function findDisplayColumn(refTable, visibleColId) {
  // If visibleCol is specified, try to find it
  if (visibleColId) {
    // visibleCol is a column ID, we need to find the column name
    // For now, just use common display columns
  }
  
  // Try common display column names
  var commonNames = ['Nom_complet', 'Nom complet', 'nom_complet', 'Name', 'name', 'Nom', 'nom', 'Label', 'label', 'Title', 'title'];
  for (var i = 0; i < commonNames.length; i++) {
    if (refTable[commonNames[i]]) return commonNames[i];
  }
  
  // Fallback: find first text column (not id, not manualSort)
  for (var col in refTable) {
    if (col !== 'id' && col !== 'manualSort' && !col.startsWith('gristHelper_')) {
      if (refTable[col] && refTable[col].length > 0 && typeof refTable[col][0] === 'string') {
        return col;
      }
    }
  }
  
  return null;
}

function lookupRefValue(refTable, refId, displayColName) {
  if (!refTable || !refTable.id || !displayColName) return null;
  
  var idx = refTable.id.indexOf(refId);
  if (idx >= 0 && refTable[displayColName]) {
    return refTable[displayColName][idx];
  }
  return null;
}


// =============================================================================
// VARIABLE CHIPS
// =============================================================================

function renderVariableChips() {
  var html = '';
  
  // Add loop syntax helper chip
  html += '<span class="var-chip" style="background:#fef3c7;color:#92400e;border:1px solid #fcd34d;" onclick="insertLoopSyntax()" title="' + t('loopExample') + '">';
  html += '🔄 ' + t('loopHint');
  html += '</span>';
  
  // Add table with loop helper chip
  html += '<span class="var-chip" style="background:#dbeafe;color:#1e40af;border:1px solid #93c5fd;" onclick="insertTableWithLoop()" title="' + t('tableLoopHint') + '">';
  html += '📊 ' + t('tableLoopBtn');
  html += '</span>';
  
  for (var i = 0; i < tableColumns.length; i++) {
    var col = tableColumns[i];
    html += '<span class="var-chip" onclick="insertVariable(\'' + sanitize(col) + '\')">';
    html += '{{' + sanitize(col) + '}}';
    html += '</span>';
  }
  document.getElementById('var-chips').innerHTML = html;
}

function insertLoopSyntax() {
  if (!editorInstance) return;
  var exampleCol = tableColumns.length > 0 ? tableColumns[0] : 'Colonne';
  var placeholder = currentLang === 'fr' ? 'Contenu répété ici...' : 'Repeated content here...';
  
  // Simple text-based loop - easier to edit
  var loopHtml = '<p>{{#each ' + exampleCol + '=Valeur}}</p>' +
    '<p>' + placeholder + '</p>' +
    '<p>{{/each}}</p>';
  
  editorInstance.selection.insertHTML(loopHtml);
  showToast(t('loopSyntax') + ' ' + (currentLang === 'fr' ? 'inséré' : 'inserted'), 'info');
}

function getUniqueValuesForColumn(colName) {
  if (!tableData || !tableData[colName]) return [];
  var values = tableData[colName];
  var unique = [];
  var seen = {};
  
  // Check if this is a date column
  var meta = columnMetadata[colName];
  var isDateColumn = meta && meta.type && (meta.type === 'Date' || meta.type === 'DateTime');
  
  // Add resolved values from tableData (current table rows)
  for (var i = 0; i < values.length; i++) {
    var val = values[i];
    if (val !== null && val !== undefined && val !== '' && !seen[val]) {
      // Format dates for display
      var displayVal = val;
      if (isDateColumn && typeof val === 'number') {
        var date = new Date(val * 1000);
        displayVal = date.toLocaleDateString(currentLang === 'fr' ? 'fr-FR' : 'en-US');
      }
      if (!seen[displayVal]) {
        seen[displayVal] = true;
        unique.push(displayVal);
      }
    }
  }
  
  // For reference columns, add ALL values from the reference table
  var meta = columnMetadata[colName];
  if (meta && meta.type) {
    var refMatch = meta.type.match(/^Ref:(.+)$/);
    if (refMatch) {
      var refTableName = refMatch[1];
      var refDisplayData = referenceDisplayValues[refTableName];
      if (refDisplayData) {
        // Add values from visibleCol (the display column configured in Grist)
        if (refDisplayData.byVisibleCol) {
          for (var refId in refDisplayData.byVisibleCol) {
            var refVal = refDisplayData.byVisibleCol[refId];
            if (refVal && !seen[refVal]) {
              seen[refVal] = true;
              unique.push(refVal);
            }
          }
        }
        // Also add values from first text column (often contains identifiers like "DUMZ 60")
        if (refDisplayData.byFirstTextCol) {
          for (var refId2 in refDisplayData.byFirstTextCol) {
            var refVal2 = refDisplayData.byFirstTextCol[refId2];
            if (refVal2 && !seen[refVal2]) {
              seen[refVal2] = true;
              unique.push(refVal2);
            }
          }
        }
      }
    }
  }
  
  return unique.sort();
}

function updateLoopValueOptions() {
  var colSelect = document.getElementById('loop-filter-col');
  var valSelect = document.getElementById('loop-filter-val-select');
  if (!colSelect || !valSelect) return;
  
  var colName = colSelect.value;
  var uniqueVals = getUniqueValuesForColumn(colName);
  
  valSelect.innerHTML = '<option value="">' + (currentLang === 'fr' ? '-- Choisir une valeur --' : '-- Choose a value --') + '</option>';
  for (var i = 0; i < uniqueVals.length; i++) {
    var opt = document.createElement('option');
    opt.value = uniqueVals[i];
    opt.textContent = uniqueVals[i];
    valSelect.appendChild(opt);
  }
}

// Cache for linked table metadata
var linkedTableMetadata = {};

async function updateLinkedTableColumns() {
  var tableSelect = document.getElementById('loop-linked-table');
  var refColSelect = document.getElementById('loop-linked-ref-col');
  var colsContainer = document.getElementById('linked-cols-container');
  var colsCheckboxes = document.getElementById('linked-cols-checkboxes');
  
  if (!tableSelect || !refColSelect || !colsCheckboxes) return;
  
  var linkedTableName = tableSelect.value;
  if (!linkedTableName) {
    refColSelect.innerHTML = '<option value="">' + (currentLang === 'fr' ? '-- Choisir la colonne Ref --' : '-- Choose Ref column --') + '</option>';
    colsCheckboxes.innerHTML = '';
    if (colsContainer) colsContainer.style.display = 'none';
    return;
  }
  
  try {
    // Fetch column metadata for the linked table
    var colData = await grist.docApi.fetchTable('_grist_Tables_column');
    var tablesData = await grist.docApi.fetchTable('_grist_Tables');
    
    // Find table ID
    var tableId = null;
    for (var i = 0; i < tablesData.id.length; i++) {
      if (tablesData.tableId[i] === linkedTableName) {
        tableId = tablesData.id[i];
        break;
      }
    }
    if (!tableId) return;
    
    // Get columns for this table
    var linkedCols = [];
    var refCols = [];
    for (var i = 0; i < colData.id.length; i++) {
      if (colData.parentId[i] === tableId) {
        var colId = colData.colId[i];
        var colType = colData.type[i];
        
        // Skip internal columns
        if (colId.startsWith('gristHelper_') || colId === 'manualSort') continue;
        
        linkedCols.push(colId);
        
        // Check if it's a Ref column pointing to the selected table
        if (colType && colType.startsWith('Ref:') && colType === 'Ref:' + selectedTable) {
          refCols.push(colId);
        }
      }
    }
    
    // Populate reference column dropdown
    refColSelect.innerHTML = '<option value="">' + (currentLang === 'fr' ? '-- Choisir la colonne Ref --' : '-- Choose Ref column --') + '</option>';
    for (var r = 0; r < refCols.length; r++) {
      var opt = document.createElement('option');
      opt.value = refCols[r];
      opt.textContent = refCols[r] + ' (→ ' + selectedTable + ')';
      refColSelect.appendChild(opt);
    }
    
    // If no ref columns found, show a message
    if (refCols.length === 0) {
      var noRefOpt = document.createElement('option');
      noRefOpt.value = '';
      noRefOpt.textContent = currentLang === 'fr' ? '⚠️ Aucune colonne Ref vers ' + selectedTable : '⚠️ No Ref column to ' + selectedTable;
      refColSelect.appendChild(noRefOpt);
    }
    
    // Populate columns checkboxes
    colsCheckboxes.innerHTML = '';
    for (var c = 0; c < linkedCols.length; c++) {
      var checked = c < 5 ? 'checked' : '';
      var label = document.createElement('label');
      label.style.cssText = 'display:block;margin-bottom:5px;cursor:pointer;';
      label.innerHTML = '<input type="checkbox" value="' + linkedCols[c] + '" ' + checked + ' style="margin-right:8px;">' + linkedCols[c];
      colsCheckboxes.appendChild(label);
    }
    
    if (colsContainer) colsContainer.style.display = 'block';
    
    // Cache metadata
    linkedTableMetadata[linkedTableName] = { columns: linkedCols, refColumns: refCols };
    
  } catch (error) {
    console.error('Error loading linked table columns:', error);
  }
}

async function updateEditLinkedTableColumns() {
  var tableSelect = document.getElementById('edit-linked-table');
  var refColSelect = document.getElementById('edit-linked-ref-col');
  
  if (!tableSelect || !refColSelect) return;
  
  var linkedTableName = tableSelect.value;
  if (!linkedTableName) {
    refColSelect.innerHTML = '<option value="">' + (currentLang === 'fr' ? '-- Choisir la colonne Ref --' : '-- Choose Ref column --') + '</option>';
    return;
  }
  
  // Keep current value if exists
  var currentValue = refColSelect.value;
  
  try {
    // Fetch column metadata for the linked table
    var colData = await grist.docApi.fetchTable('_grist_Tables_column');
    var tablesData = await grist.docApi.fetchTable('_grist_Tables');
    
    // Find table ID
    var tableId = null;
    for (var i = 0; i < tablesData.id.length; i++) {
      if (tablesData.tableId[i] === linkedTableName) {
        tableId = tablesData.id[i];
        break;
      }
    }
    if (!tableId) return;
    
    // Get ref columns for this table
    var refCols = [];
    for (var i = 0; i < colData.id.length; i++) {
      if (colData.parentId[i] === tableId) {
        var colId = colData.colId[i];
        var colType = colData.type[i];
        
        // Check if it's a Ref column pointing to the selected table
        if (colType && colType.startsWith('Ref:') && colType === 'Ref:' + selectedTable) {
          refCols.push(colId);
        }
      }
    }
    
    // Populate reference column dropdown
    refColSelect.innerHTML = '<option value="">' + (currentLang === 'fr' ? '-- Choisir la colonne Ref --' : '-- Choose Ref column --') + '</option>';
    for (var r = 0; r < refCols.length; r++) {
      var opt = document.createElement('option');
      opt.value = refCols[r];
      opt.textContent = refCols[r] + ' (→ ' + selectedTable + ')';
      if (refCols[r] === currentValue) opt.selected = true;
      refColSelect.appendChild(opt);
    }
    
    // If no ref columns found, show a message
    if (refCols.length === 0) {
      var noRefOpt = document.createElement('option');
      noRefOpt.value = '';
      noRefOpt.textContent = currentLang === 'fr' ? '⚠️ Aucune colonne Ref vers ' + selectedTable : '⚠️ No Ref column to ' + selectedTable;
      refColSelect.appendChild(noRefOpt);
    }
    
  } catch (error) {
    console.error('Error loading linked table columns:', error);
  }
}

function updateEditLoopValueOptions() {
  var colSelect = document.getElementById('edit-loop-filter-col');
  var valSelect = document.getElementById('edit-loop-filter-val-select');
  if (!colSelect || !valSelect) return;
  
  var colName = colSelect.value;
  var uniqueVals = getUniqueValuesForColumn(colName);
  
  valSelect.innerHTML = '<option value="">' + (currentLang === 'fr' ? '-- Choisir une valeur --' : '-- Choose a value --') + '</option>';
  for (var i = 0; i < uniqueVals.length; i++) {
    var opt = document.createElement('option');
    opt.value = uniqueVals[i];
    opt.textContent = uniqueVals[i];
    valSelect.appendChild(opt);
  }
}

function insertTableWithLoop() {
  if (!editorInstance) return;
  
  // Build column selector options
  var colOptions = '';
  for (var i = 0; i < tableColumns.length; i++) {
    colOptions += '<option value="' + tableColumns[i] + '">' + tableColumns[i] + '</option>';
  }
  
  // Build view selector options
  var viewOptions = '<option value="">' + (currentLang === 'fr' ? '-- Choisir une vue --' : '-- Choose a view --') + '</option>';
  for (var v = 0; v < availableViews.length; v++) {
    var viewName = availableViews[v].name;
    var viewId = availableViews[v].id;
    var hasFilters = availableViews[v].filters ? ' 🔍' : '';
    viewOptions += '<option value="' + viewId + '">' + viewName + hasFilters + '</option>';
  }
  
  var viewLinkedHelp = currentLang === 'fr' 
    ? '💡 Utilise les lignes visibles via "Sélectionner par" (panneau Grist à droite)'
    : '💡 Uses visible rows via "Select By" (Grist panel on the right)';
  
  var viewSelectHelp = currentLang === 'fr'
    ? '💡 Sélectionnez une vue existante pour utiliser ses filtres'
    : '💡 Select an existing view to use its filters';
  
  var linkedTableHelp = currentLang === 'fr'
    ? '💡 Affiche les lignes d\'une autre table liées à l\'enregistrement courant (ex: Facture_details liée à Facture)'
    : '💡 Shows rows from another table linked to the current record (e.g., Invoice_details linked to Invoice)';
  
  // Build table selector options (all tables except current)
  var tableOptions = '<option value="">' + (currentLang === 'fr' ? '-- Choisir une table --' : '-- Choose a table --') + '</option>';
  for (var t = 0; t < allTables.length; t++) {
    if (allTables[t] !== selectedTable) {
      tableOptions += '<option value="' + allTables[t] + '">' + allTables[t] + '</option>';
    }
  }
  
  var formHtml = '<div style="text-align:left;">' +
    '<div style="margin-bottom:15px;">' +
    '<label style="display:block;margin-bottom:8px;font-weight:600;">' + (currentLang === 'fr' ? 'Type de tableau :' : 'Table type:') + '</label>' +
    '<label style="display:block;margin-bottom:5px;cursor:pointer;">' +
    '<input type="radio" name="loop-type" value="view" checked style="margin-right:8px;">' +
    (currentLang === 'fr' ? 'Lié à la vue courante' : 'Linked to current view') + '</label>' +
    '<p style="margin:0 0 10px 24px;font-size:0.85em;color:#6b7280;">' + viewLinkedHelp + '</p>' +
    '<label style="display:block;margin-bottom:5px;cursor:pointer;">' +
    '<input type="radio" name="loop-type" value="viewselect" style="margin-right:8px;">' +
    (currentLang === 'fr' ? 'Lié à une vue filtrée' : 'Linked to a filtered view') + '</label>' +
    '<p style="margin:0 0 10px 24px;font-size:0.85em;color:#6b7280;">' + viewSelectHelp + '</p>' +
    '<label style="display:block;margin-bottom:5px;cursor:pointer;">' +
    '<input type="radio" name="loop-type" value="linkedtable" style="margin-right:8px;">' +
    (currentLang === 'fr' ? 'Lié à une table externe' : 'Linked to external table') + '</label>' +
    '<p style="margin:0 0 10px 24px;font-size:0.85em;color:#6b7280;">' + linkedTableHelp + '</p>' +
    '<label style="display:block;margin-bottom:5px;cursor:pointer;">' +
    '<input type="radio" name="loop-type" value="filter" style="margin-right:8px;">' +
    (currentLang === 'fr' ? 'Avec filtre manuel' : 'With manual filter') + '</label>' +
    '</div>' +
    '<div id="view-select-options" style="display:none;border:1px solid #e5e7eb;padding:10px;border-radius:6px;margin-bottom:10px;background:#f0fdf4;">' +
    '<label style="display:block;margin-bottom:5px;font-weight:600;">' + (currentLang === 'fr' ? 'Vue à utiliser :' : 'View to use:') + '</label>' +
    '<select id="loop-view-select" style="width:100%;padding:8px;border:1px solid #ccc;border-radius:4px;">' + viewOptions + '</select>' +
    '</div>' +
    '<div id="linked-table-options" style="display:none;border:1px solid #e5e7eb;padding:10px;border-radius:6px;margin-bottom:10px;background:#fef3c7;">' +
    '<div style="margin-bottom:10px;">' +
    '<label style="display:block;margin-bottom:5px;font-weight:600;">' + (currentLang === 'fr' ? 'Table à afficher :' : 'Table to display:') + '</label>' +
    '<select id="loop-linked-table" onchange="updateLinkedTableColumns()" style="width:100%;padding:8px;border:1px solid #ccc;border-radius:4px;">' + tableOptions + '</select>' +
    '</div>' +
    '<div style="margin-bottom:10px;">' +
    '<label style="display:block;margin-bottom:5px;font-weight:600;">' + (currentLang === 'fr' ? 'Colonne de référence (vers ' + selectedTable + ') :' : 'Reference column (to ' + selectedTable + '):') + '</label>' +
    '<select id="loop-linked-ref-col" style="width:100%;padding:8px;border:1px solid #ccc;border-radius:4px;">' +
    '<option value="">' + (currentLang === 'fr' ? '-- Choisir la colonne Ref --' : '-- Choose Ref column --') + '</option>' +
    '</select>' +
    '</div>' +
    '<div id="linked-cols-container" style="display:none;">' +
    '<label style="display:block;margin-bottom:5px;font-weight:600;">' + (currentLang === 'fr' ? 'Colonnes à afficher :' : 'Columns to display:') + '</label>' +
    '<div id="linked-cols-checkboxes" style="max-height:120px;overflow-y:auto;border:1px solid #eee;padding:8px;border-radius:4px;background:white;"></div>' +
    '</div>' +
    '</div>' +
    '<div id="filter-options" style="display:none;border:1px solid #e5e7eb;padding:10px;border-radius:6px;margin-bottom:10px;background:#f9fafb;">' +
    '<div style="margin-bottom:10px;">' +
    '<label style="display:block;margin-bottom:5px;font-weight:600;">' + (currentLang === 'fr' ? 'Colonne à filtrer :' : 'Column to filter:') + '</label>' +
    '<select id="loop-filter-col" onchange="updateLoopValueOptions()" style="width:100%;padding:8px;border:1px solid #ccc;border-radius:4px;">' + colOptions + '</select>' +
    '</div>' +
    '<div style="margin-bottom:10px;">' +
    '<label style="display:block;margin-bottom:5px;font-weight:600;">' + (currentLang === 'fr' ? 'Valeur à rechercher :' : 'Value to search:') + '</label>' +
    '<select id="loop-filter-val-select" style="width:100%;padding:8px;border:1px solid #ccc;border-radius:4px;margin-bottom:5px;">' +
    '<option value="">' + (currentLang === 'fr' ? '-- Choisir une valeur --' : '-- Choose a value --') + '</option>' +
    '</select>' +
    '<input type="text" id="loop-filter-val" placeholder="' + (currentLang === 'fr' ? 'Ou saisir manuellement...' : 'Or type manually...') + '" style="width:100%;padding:8px;border:1px solid #ccc;border-radius:4px;box-sizing:border-box;">' +
    '</div>' +
    '</div>' +
    '<div style="margin-bottom:10px;">' +
    '<label style="display:block;margin-bottom:5px;font-weight:600;">' + (currentLang === 'fr' ? 'Colonnes à afficher :' : 'Columns to display:') + '</label>' +
    '<div id="loop-cols-checkboxes" style="max-height:150px;overflow-y:auto;border:1px solid #eee;padding:8px;border-radius:4px;">';
  
  for (var j = 0; j < tableColumns.length; j++) {
    var checked = j < 5 ? 'checked' : '';
    formHtml += '<label style="display:block;margin-bottom:5px;cursor:pointer;">' +
      '<input type="checkbox" value="' + tableColumns[j] + '" ' + checked + ' style="margin-right:8px;">' +
      tableColumns[j] + '</label>';
  }
  
  formHtml += '</div></div></div>';
  
  // Show modal and initialize value dropdown + radio button handlers
  setTimeout(function() { 
    updateLoopValueOptions();
    // Add event listeners for radio buttons
    var radios = document.querySelectorAll('input[name="loop-type"]');
    radios.forEach(function(radio) {
      radio.addEventListener('change', function() {
        var filterOptions = document.getElementById('filter-options');
        var viewSelectOptions = document.getElementById('view-select-options');
        var linkedTableOptions = document.getElementById('linked-table-options');
        var loopColsSection = document.getElementById('loop-cols-checkboxes');
        if (filterOptions) {
          filterOptions.style.display = this.value === 'filter' ? 'block' : 'none';
        }
        if (viewSelectOptions) {
          viewSelectOptions.style.display = this.value === 'viewselect' ? 'block' : 'none';
        }
        if (linkedTableOptions) {
          linkedTableOptions.style.display = this.value === 'linkedtable' ? 'block' : 'none';
        }
        // Hide main columns section when linkedtable is selected (it has its own)
        if (loopColsSection && loopColsSection.parentElement) {
          loopColsSection.parentElement.style.display = this.value === 'linkedtable' ? 'none' : 'block';
        }
      });
    });
  }, 100);
  
  showModal(currentLang === 'fr' ? '📊 Tableau avec boucle' : '📊 Table with loop', formHtml).then(function(confirmed) {
    if (!confirmed) return;
    
    // Check which type is selected
    var loopType = document.querySelector('input[name="loop-type"]:checked');
    var loopTypeValue = loopType ? loopType.value : 'view';
    
    var filterCol = '';
    var filterVal = '';
    var selectedViewId = '';
    
    var linkedTableName = '';
    var linkedRefCol = '';
    
    if (loopTypeValue === 'filter') {
      filterCol = document.getElementById('loop-filter-col').value;
      // Use dropdown value if selected, otherwise use text input
      var filterValSelect = document.getElementById('loop-filter-val-select');
      var filterValInput = document.getElementById('loop-filter-val');
      filterVal = (filterValSelect && filterValSelect.value) || (filterValInput && filterValInput.value) || (currentLang === 'fr' ? 'Valeur' : 'Value');
    } else if (loopTypeValue === 'viewselect') {
      var viewSelect = document.getElementById('loop-view-select');
      selectedViewId = viewSelect ? viewSelect.value : '';
    } else if (loopTypeValue === 'linkedtable') {
      var linkedTableSelect = document.getElementById('loop-linked-table');
      var linkedRefColSelect = document.getElementById('loop-linked-ref-col');
      linkedTableName = linkedTableSelect ? linkedTableSelect.value : '';
      linkedRefCol = linkedRefColSelect ? linkedRefColSelect.value : '';
    }
    
    // Get selected columns (from linked table if linkedtable type)
    var checkboxes;
    var selectedCols = [];
    if (loopTypeValue === 'linkedtable') {
      checkboxes = document.querySelectorAll('#linked-cols-checkboxes input[type="checkbox"]:checked');
    } else {
      checkboxes = document.querySelectorAll('#loop-cols-checkboxes input[type="checkbox"]:checked');
    }
    checkboxes.forEach(function(cb) { selectedCols.push(cb.value); });
    
    if (selectedCols.length === 0) {
      if (loopTypeValue === 'linkedtable' && linkedTableMetadata[linkedTableName]) {
        selectedCols = linkedTableMetadata[linkedTableName].columns.slice(0, 5);
      } else {
        selectedCols = tableColumns.slice(0, 5);
      }
    }
    
    // Build table HTML
    var headerCells = '';
    var dataCells = '';
    for (var k = 0; k < selectedCols.length; k++) {
      headerCells += '<th style="border:1px solid #ccc;padding:8px;background:#f3f4f6;">' + selectedCols[k] + '</th>';
      dataCells += '<td style="border:1px solid #ccc;padding:8px;">{{' + selectedCols[k] + '}}</td>';
    }
    
    var tableHtml;
    if (loopTypeValue === 'view') {
      // View-linked table: uses <!--LOOP:*--> to show all rows from the current view
      tableHtml = '<table style="border-collapse:collapse;width:100%;margin:10px 0;">' +
        '<thead><tr>' + headerCells + '</tr></thead>' +
        '<tbody>' +
        '<!--LOOP:*-->' +
        '<tr>' + dataCells + '</tr>' +
        '<!--/LOOP-->' +
        '</tbody>' +
        '</table>';
    } else if (loopTypeValue === 'viewselect' && selectedViewId) {
      // View-select table: uses <!--LOOP:VIEW:viewId--> to show rows from selected view
      tableHtml = '<table style="border-collapse:collapse;width:100%;margin:10px 0;">' +
        '<thead><tr>' + headerCells + '</tr></thead>' +
        '<tbody>' +
        '<!--LOOP:VIEW:' + selectedViewId + '-->' +
        '<tr>' + dataCells + '</tr>' +
        '<!--/LOOP-->' +
        '</tbody>' +
        '</table>';
    } else if (loopTypeValue === 'linkedtable' && linkedTableName && linkedRefCol) {
      // Linked table: uses <!--LOOP:TABLE:tableName:refColumn--> to show rows from linked table
      tableHtml = '<table style="border-collapse:collapse;width:100%;margin:10px 0;">' +
        '<thead><tr>' + headerCells + '</tr></thead>' +
        '<tbody>' +
        '<!--LOOP:TABLE:' + linkedTableName + ':' + linkedRefCol + '-->' +
        '<tr>' + dataCells + '</tr>' +
        '<!--/LOOP-->' +
        '</tbody>' +
        '</table>';
    } else {
      // Filtered table
      tableHtml = '<table style="border-collapse:collapse;width:100%;margin:10px 0;">' +
        '<thead><tr>' + headerCells + '</tr></thead>' +
        '<tbody>' +
        '<!--LOOP:' + filterCol + '=' + filterVal + '-->' +
        '<tr>' + dataCells + '</tr>' +
        '<!--/LOOP-->' +
        '</tbody>' +
        '</table>';
    }
    
    editorInstance.selection.insertHTML(tableHtml);
    showToast(currentLang === 'fr' ? 'Tableau avec boucle inséré' : 'Table with loop inserted', 'info');
  });
}

function insertVariable(colName) {
  if (!editorInstance) return;
  var varHtml = '<span style="background:#f3e8ff;color:#7c3aed;padding:2px 6px;border-radius:4px;font-weight:600;" contenteditable="false">{{' + colName + '}}</span>&nbsp;';
  editorInstance.selection.insertHTML(varHtml);
  showToast('{{' + colName + '}} inséré', 'info');
}

// Edit existing loop in a table
function editTableLoop(tableElement) {
  if (!editorInstance || !tableElement) return;
  
  // Find the loop comment in the table
  var tbody = tableElement.querySelector('tbody');
  if (!tbody) return;
  
  var loopComment = null;
  var currentFilterCol = '';
  var currentFilterVal = '';
  
  // Search for loop comment in tbody
  var isViewLinked = false;
  var isViewSelect = false;
  var isLinkedTable = false;
  var currentViewId = '';
  var currentLinkedTable = '';
  var currentLinkedRefCol = '';
  for (var i = 0; i < tbody.childNodes.length; i++) {
    var node = tbody.childNodes[i];
    if (node.nodeType === 8) { // Comment node
      if (node.textContent === 'LOOP:*') {
        loopComment = node;
        isViewLinked = true;
        break;
      }
      var viewMatch = node.textContent.match(/^LOOP:VIEW:(\d+)$/);
      if (viewMatch) {
        loopComment = node;
        isViewSelect = true;
        currentViewId = viewMatch[1];
        break;
      }
      var tableMatch = node.textContent.match(/^LOOP:TABLE:([^:]+):(.+)$/);
      if (tableMatch) {
        loopComment = node;
        isLinkedTable = true;
        currentLinkedTable = tableMatch[1];
        currentLinkedRefCol = tableMatch[2];
        break;
      }
      var match = node.textContent.match(/^LOOP:([^=]+)=(.*)$/);
      if (match) {
        loopComment = node;
        currentFilterCol = match[1];
        currentFilterVal = match[2];
        break;
      }
    }
  }
  
  if (!loopComment) {
    showToast(currentLang === 'fr' ? 'Aucune boucle trouvée dans ce tableau' : 'No loop found in this table', 'error');
    return;
  }
  
  // For linked table loops, show full editing modal
  if (isLinkedTable) {
    // Build table selector options
    var linkedTableOptions = '<option value="">' + (currentLang === 'fr' ? '-- Choisir une table --' : '-- Choose a table --') + '</option>';
    for (var t = 0; t < allTables.length; t++) {
      if (allTables[t] !== selectedTable) {
        var selectedOpt = allTables[t] === currentLinkedTable ? 'selected' : '';
        linkedTableOptions += '<option value="' + allTables[t] + '" ' + selectedOpt + '>' + allTables[t] + '</option>';
      }
    }
    
    var linkedFormHtml = '<div style="text-align:left;">' +
      '<div style="margin-bottom:15px;">' +
      '<label style="display:block;margin-bottom:5px;font-weight:600;">' + (currentLang === 'fr' ? 'Table liée :' : 'Linked table:') + '</label>' +
      '<select id="edit-linked-table" onchange="updateEditLinkedTableColumns()" style="width:100%;padding:8px;border:1px solid #ccc;border-radius:4px;">' + linkedTableOptions + '</select>' +
      '</div>' +
      '<div style="margin-bottom:15px;">' +
      '<label style="display:block;margin-bottom:5px;font-weight:600;">' + (currentLang === 'fr' ? 'Colonne de référence (vers ' + selectedTable + ') :' : 'Reference column (to ' + selectedTable + '):') + '</label>' +
      '<select id="edit-linked-ref-col" style="width:100%;padding:8px;border:1px solid #ccc;border-radius:4px;">' +
      '<option value="' + currentLinkedRefCol + '" selected>' + currentLinkedRefCol + '</option>' +
      '</select>' +
      '</div>' +
      '</div>';
    
    // Show modal and handle confirmation
    setTimeout(function() {
      // Load columns for the current linked table
      updateEditLinkedTableColumns();
    }, 100);
    
    showModal(currentLang === 'fr' ? '🔗 Modifier le tableau lié' : '🔗 Edit linked table', linkedFormHtml).then(function(confirmed) {
      if (!confirmed) return;
      
      var newLinkedTable = document.getElementById('edit-linked-table').value;
      var newRefCol = document.getElementById('edit-linked-ref-col').value;
      
      if (!newLinkedTable || !newRefCol) {
        showToast(currentLang === 'fr' ? 'Veuillez sélectionner une table et une colonne de référence' : 'Please select a table and reference column', 'error');
        return;
      }
      
      // Update the loop comment
      loopComment.textContent = 'LOOP:TABLE:' + newLinkedTable + ':' + newRefCol;
      
      // Update editor content
      var updatedHtml = tableElement.outerHTML;
      
      showToast(currentLang === 'fr' ? 'Boucle mise à jour' : 'Loop updated', 'success');
    });
    return;
  }
  
  // Build column selector options
  var colOptions = '';
  for (var i = 0; i < tableColumns.length; i++) {
    var selected = tableColumns[i] === currentFilterCol ? 'selected' : '';
    colOptions += '<option value="' + tableColumns[i] + '" ' + selected + '>' + tableColumns[i] + '</option>';
  }
  
  // Build view selector options
  var viewOptions = '<option value="">' + (currentLang === 'fr' ? '-- Choisir une vue --' : '-- Choose a view --') + '</option>';
  for (var v = 0; v < availableViews.length; v++) {
    var viewName = availableViews[v].name;
    var viewId = availableViews[v].id;
    var hasFilters = availableViews[v].filters ? ' 🔍' : '';
    var selectedView = String(viewId) === currentViewId ? 'selected' : '';
    viewOptions += '<option value="' + viewId + '" ' + selectedView + '>' + viewName + hasFilters + '</option>';
  }
  
  // Build value options for current column
  var uniqueVals = getUniqueValuesForColumn(currentFilterCol || tableColumns[0]);
  var valOptions = '<option value="">' + (currentLang === 'fr' ? '-- Choisir une valeur --' : '-- Choose a value --') + '</option>';
  for (var j = 0; j < uniqueVals.length; j++) {
    var selVal = uniqueVals[j] === currentFilterVal ? 'selected' : '';
    valOptions += '<option value="' + uniqueVals[j] + '" ' + selVal + '>' + uniqueVals[j] + '</option>';
  }
  
  // Determine which radio should be checked
  var isFilterChecked = !isViewLinked && !isViewSelect;
  
  var formHtml = '<div style="text-align:left;">' +
    '<div style="margin-bottom:15px;">' +
    '<label style="display:block;margin-bottom:8px;font-weight:600;">' + (currentLang === 'fr' ? 'Type de tableau :' : 'Table type:') + '</label>' +
    '<label style="display:block;margin-bottom:5px;cursor:pointer;">' +
    '<input type="radio" name="edit-loop-type" value="view" ' + (isViewLinked ? 'checked' : '') + ' style="margin-right:8px;">' +
    (currentLang === 'fr' ? 'Lié à la vue courante' : 'Linked to current view') + '</label>' +
    '<label style="display:block;margin-bottom:5px;cursor:pointer;">' +
    '<input type="radio" name="edit-loop-type" value="viewselect" ' + (isViewSelect ? 'checked' : '') + ' style="margin-right:8px;">' +
    (currentLang === 'fr' ? 'Lié à une vue filtrée' : 'Linked to a filtered view') + '</label>' +
    '<label style="display:block;margin-bottom:5px;cursor:pointer;">' +
    '<input type="radio" name="edit-loop-type" value="filter" ' + (isFilterChecked ? 'checked' : '') + ' style="margin-right:8px;">' +
    (currentLang === 'fr' ? 'Avec filtre manuel' : 'With manual filter') + '</label>' +
    '</div>' +
    '<div id="edit-view-select-options" style="' + (isViewSelect ? '' : 'display:none;') + 'border:1px solid #e5e7eb;padding:10px;border-radius:6px;margin-bottom:10px;background:#f0fdf4;">' +
    '<label style="display:block;margin-bottom:5px;font-weight:600;">' + (currentLang === 'fr' ? 'Vue à utiliser :' : 'View to use:') + '</label>' +
    '<select id="edit-loop-view-select" style="width:100%;padding:8px;border:1px solid #ccc;border-radius:4px;">' + viewOptions + '</select>' +
    '</div>' +
    '<div id="edit-filter-options" style="' + (isFilterChecked ? '' : 'display:none;') + 'border:1px solid #e5e7eb;padding:10px;border-radius:6px;background:#f9fafb;">' +
    '<div style="margin-bottom:10px;">' +
    '<label style="display:block;margin-bottom:5px;font-weight:600;">' + (currentLang === 'fr' ? 'Colonne à filtrer :' : 'Column to filter:') + '</label>' +
    '<select id="edit-loop-filter-col" onchange="updateEditLoopValueOptions()" style="width:100%;padding:8px;border:1px solid #ccc;border-radius:4px;">' + colOptions + '</select>' +
    '</div>' +
    '<div style="margin-bottom:10px;">' +
    '<label style="display:block;margin-bottom:5px;font-weight:600;">' + (currentLang === 'fr' ? 'Valeur à rechercher :' : 'Value to search:') + '</label>' +
    '<select id="edit-loop-filter-val-select" style="width:100%;padding:8px;border:1px solid #ccc;border-radius:4px;margin-bottom:5px;">' + valOptions + '</select>' +
    '<input type="text" id="edit-loop-filter-val" value="' + currentFilterVal + '" placeholder="' + (currentLang === 'fr' ? 'Ou saisir manuellement...' : 'Or type manually...') + '" style="width:100%;padding:8px;border:1px solid #ccc;border-radius:4px;box-sizing:border-box;">' +
    '</div>' +
    '</div>' +
    '</div>';
  
  // Add event listeners for radio buttons after modal opens
  setTimeout(function() {
    var radios = document.querySelectorAll('input[name="edit-loop-type"]');
    radios.forEach(function(radio) {
      radio.addEventListener('change', function() {
        var filterOptions = document.getElementById('edit-filter-options');
        var viewSelectOptions = document.getElementById('edit-view-select-options');
        if (filterOptions) {
          filterOptions.style.display = this.value === 'filter' ? 'block' : 'none';
        }
        if (viewSelectOptions) {
          viewSelectOptions.style.display = this.value === 'viewselect' ? 'block' : 'none';
        }
      });
    });
  }, 100);
  
  showModal(currentLang === 'fr' ? '✏️ Modifier la boucle' : '✏️ Edit loop', formHtml).then(function(confirmed) {
    if (!confirmed) return;
    
    var loopType = document.querySelector('input[name="edit-loop-type"]:checked');
    var loopTypeValue = loopType ? loopType.value : 'view';
    
    if (loopTypeValue === 'view') {
      loopComment.textContent = 'LOOP:*';
    } else if (loopTypeValue === 'viewselect') {
      var viewSelect = document.getElementById('edit-loop-view-select');
      var selectedViewId = viewSelect ? viewSelect.value : '';
      if (selectedViewId) {
        loopComment.textContent = 'LOOP:VIEW:' + selectedViewId;
      } else {
        loopComment.textContent = 'LOOP:*';
      }
    } else {
      var newFilterCol = document.getElementById('edit-loop-filter-col').value;
      var newFilterValSelect = document.getElementById('edit-loop-filter-val-select');
      var newFilterValInput = document.getElementById('edit-loop-filter-val');
      var newFilterVal = (newFilterValSelect && newFilterValSelect.value) || (newFilterValInput && newFilterValInput.value) || currentFilterVal;
      loopComment.textContent = 'LOOP:' + newFilterCol + '=' + newFilterVal;
    }
    
    showToast(currentLang === 'fr' ? 'Boucle modifiée !' : 'Loop updated!', 'success');
    scheduleAutoSave();
  });
}

// Remove loop edit button if exists
function removeLoopEditButton() {
  var existing = document.getElementById('loop-edit-btn');
  if (existing) existing.remove();
}

// Show loop edit button near a table
function showLoopEditButton(tableElement, event) {
  removeLoopEditButton();
  
  var btn = document.createElement('button');
  btn.id = 'loop-edit-btn';
  btn.innerHTML = '✏️ ' + (currentLang === 'fr' ? 'Modifier la boucle' : 'Edit loop');
  btn.style.cssText = 'position:absolute;z-index:9999;background:#8b5cf6;color:white;border:none;padding:6px 12px;border-radius:6px;font-size:12px;cursor:pointer;box-shadow:0 2px 8px rgba(0,0,0,0.2);';
  btn.onclick = function(e) {
    e.stopPropagation();
    editTableLoop(tableElement);
    removeLoopEditButton();
  };
  
  // Position near the click
  var editorArea = document.querySelector('.jodit-wysiwyg');
  if (editorArea) {
    var rect = editorArea.getBoundingClientRect();
    btn.style.left = (event.clientX - rect.left + editorArea.scrollLeft) + 'px';
    btn.style.top = (event.clientY - rect.top + editorArea.scrollTop - 40) + 'px';
    editorArea.style.position = 'relative';
    editorArea.appendChild(btn);
    
    // Auto-hide after 5 seconds
    setTimeout(removeLoopEditButton, 5000);
  }
}

// =============================================================================
// JODIT EDITOR
// =============================================================================

function initEditor() {
  editorInstance = Jodit.make('#editor-container', {
    language: currentLang,
    minHeight: 500,
    placeholder: currentLang === 'fr' ? 'Commencez à écrire votre document ici...' : 'Start writing your document here...',
    allowResizeY: true,
    toolbarAdaptive: false,
    askBeforePasteHTML: false,
    askBeforePasteFromWord: false,
    defaultActionOnPaste: 'insert_clear_html',
    controls: {
      pagebreak: {
        name: 'pagebreak',
        iconURL: '',
        tooltip: currentLang === 'fr' ? 'Saut de page' : 'Page break',
        exec: function(editor) {
          editor.selection.insertHTML(
            '<div class="page-break-marker" contenteditable="false" style="page-break-after:always;">' +
            '✂ ' + (currentLang === 'fr' ? 'Saut de page' : 'Page break') +
            '</div>'
          );
        }
      },
      verticaltext: {
        name: 'verticaltext',
        iconURL: '',
        tooltip: currentLang === 'fr' ? 'Texte vertical (cliquer pour changer le sens)' : 'Vertical text (click to change direction)',
        exec: function(editor) {
          var selection = editor.selection;
          var current = selection.current();
          if (current) {
            // Find the closest cell (td or th)
            var cell = current.closest ? current.closest('td, th') : null;
            if (!cell) {
              var el = current;
              while (el && el.tagName !== 'TD' && el.tagName !== 'TH') {
                el = el.parentElement;
              }
              cell = el;
            }
            if (cell) {
              // Cycle through: normal -> vertical-rl (bottom to top) -> vertical-lr (top to bottom) -> normal
              var currentMode = cell.style.writingMode;
              if (currentMode === 'vertical-rl') {
                // Switch to vertical-lr (top to bottom, text upright)
                cell.style.writingMode = 'vertical-lr';
                cell.style.textOrientation = 'mixed';
                cell.style.transform = 'rotate(180deg)';
                cell.style.whiteSpace = 'nowrap';
              } else if (currentMode === 'vertical-lr') {
                // Back to normal
                cell.style.writingMode = '';
                cell.style.textOrientation = '';
                cell.style.transform = '';
                cell.style.whiteSpace = '';
              } else {
                // Start with vertical-rl (bottom to top)
                cell.style.writingMode = 'vertical-rl';
                cell.style.textOrientation = 'mixed';
                cell.style.transform = '';
                cell.style.whiteSpace = 'nowrap';
              }
            } else {
              // Not in a table cell, wrap selection in a span
              var html = selection.html;
              if (html) {
                selection.insertHTML('<span style="writing-mode:vertical-rl;text-orientation:mixed;display:inline-block;">' + html + '</span>');
              }
            }
          }
        }
      }
    },
    buttons: [
      'bold', 'italic', 'underline', 'strikethrough', '|',
      'font', 'fontsize', '|',
      'brush', '|',
      'paragraph', '|',
      'ul', 'ol', '|',
      'outdent', 'indent', '|',
      'align', 'verticaltext', '|',
      'table', '|',
      'link', 'image', '|',
      'hr', 'pagebreak', '|',
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
    showXPathInStatusbar: false,
    events: {
      change: function() {
        scheduleAutoSave();
      },
      click: function(e) {
        // Check if clicked on a table with loop
        var target = e.target;
        var table = target.closest ? target.closest('table') : null;
        if (!table) {
          // Try parent elements
          var el = target;
          while (el && el.tagName !== 'TABLE') {
            el = el.parentElement;
          }
          table = el;
        }
        
        if (table) {
          // Check if this table has a loop comment
          var tbody = table.querySelector('tbody');
          if (tbody) {
            var hasLoop = false;
            for (var i = 0; i < tbody.childNodes.length; i++) {
              var node = tbody.childNodes[i];
              if (node.nodeType === 8 && node.textContent.match(/^LOOP:/)) {
                hasLoop = true;
                break;
              }
            }
            if (hasLoop) {
              showLoopEditButton(table, e);
              return;
            }
          }
        }
        removeLoopEditButton();
      }
    }
  });
  
}

// =============================================================================
// PAGE FORMAT VISUAL
// =============================================================================

function setEditorPageFormat(format) {
  var wrapper = document.getElementById('editor-page-wrapper');
  if (!wrapper) return;
  wrapper.className = '';
  if (format === 'a4') {
    wrapper.className = 'editor-page-a4';
  } else if (format === 'letter') {
    wrapper.className = 'editor-page-letter';
  } else {
    wrapper.className = 'editor-page-free';
  }
  // Sync with PDF page size selector if it exists
  var pdfPageSize = document.getElementById('pdf-page-size');
  if (pdfPageSize && format !== 'free') {
    pdfPageSize.value = format;
  }
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

// --- Multi-template management ---

var currentTemplateName = '';

async function getTemplateIndex() {
  // Returns { templates: [ { name: "...", html: "..." }, ... ] }
  try {
    var idx = await grist.widgetApi.getOption('templateIndex');
    console.log('getTemplateIndex from Grist:', idx);
    if (idx && Array.isArray(idx.templates)) return idx;
  } catch (e) {
    console.warn('Error getting templateIndex from Grist:', e);
  }
  // Fallback to localStorage
  try {
    var local = localStorage.getItem(TEMPLATE_STORAGE_KEY + 'index');
    console.log('getTemplateIndex from localStorage:', local);
    if (local) {
      var parsed = JSON.parse(local);
      if (parsed && Array.isArray(parsed.templates)) return parsed;
    }
  } catch (e) {
    console.warn('Error getting templateIndex from localStorage:', e);
  }
  console.log('getTemplateIndex: returning empty');
  return { templates: [] };
}

async function saveTemplateIndex(index) {
  console.log('Saving template index:', JSON.stringify(index));
  try {
    await grist.widgetApi.setOption('templateIndex', index);
    console.log('Template index saved to Grist options');
  } catch (e) {
    console.warn('Could not save template index to Grist:', e);
  }
  try {
    localStorage.setItem(TEMPLATE_STORAGE_KEY + 'index', JSON.stringify(index));
    console.log('Template index saved to localStorage');
  } catch (e) {
    console.warn('Could not save to localStorage:', e);
  }
}

async function refreshTemplateList() {
  var select = document.getElementById('template-list');
  var saveTargetSelect = document.getElementById('save-target-select');
  if (!select) return;
  var index = await getTemplateIndex();
  select.innerHTML = '<option value="">' + t('templateSelectDefault') + '</option>';
  
  // Also update save target dropdown
  if (saveTargetSelect) {
    saveTargetSelect.innerHTML = '<option value="new">' + (currentLang === 'fr' ? '➕ Nouveau modèle' : '➕ New template') + '</option>';
  }
  
  for (var i = 0; i < index.templates.length; i++) {
    var opt = document.createElement('option');
    opt.value = index.templates[i].name;
    opt.textContent = index.templates[i].name;
    if (index.templates[i].name === currentTemplateName) opt.selected = true;
    select.appendChild(opt);
    
    // Add to save target dropdown too
    if (saveTargetSelect) {
      var saveOpt = document.createElement('option');
      saveOpt.value = index.templates[i].name;
      saveOpt.textContent = '📝 ' + index.templates[i].name;
      saveTargetSelect.appendChild(saveOpt);
    }
  }
  // Show/hide delete button
  var delBtn = document.getElementById('btn-delete-template');
  if (delBtn) delBtn.style.display = select.value ? '' : 'none';
}

function onTemplateSelectChange() {
  var select = document.getElementById('template-list');
  var nameInput = document.getElementById('template-name-input');
  var delBtn = document.getElementById('btn-delete-template');
  if (!select) return;
  if (select.value) {
    // Load selected template
    loadTemplateByName(select.value);
    if (nameInput) nameInput.value = select.value;
    if (delBtn) delBtn.style.display = '';
  } else {
    // New template mode — clear editor
    if (nameInput) nameInput.value = '';
    currentTemplateName = '';
    if (delBtn) delBtn.style.display = 'none';
    if (editorInstance) {
      editorInstance.value = '';
      templateHtml = '';
    }
  }
}

async function loadTemplateByName(name) {
  var index = await getTemplateIndex();
  for (var i = 0; i < index.templates.length; i++) {
    if (index.templates[i].name === name) {
      setEditorHtml(index.templates[i].html);
      templateHtml = index.templates[i].html;
      currentTemplateName = name;
      showToast(t('templateLoaded'), 'info');
      return;
    }
  }
}

async function saveTemplate() {
  if (!editorInstance) return;
  templateHtml = getEditorHtml();
  if (!selectedTable) return;

  var nameInput = document.getElementById('template-name-input');
  var saveTargetSelect = document.getElementById('save-target-select');
  var saveTarget = saveTargetSelect ? saveTargetSelect.value : 'new';
  
  var name;
  if (saveTarget === 'new') {
    // Creating new template - use name input
    name = (nameInput ? nameInput.value.trim() : '');
    if (!name) {
      // Prompt for name
      name = prompt(t('templateName'), t('templateNamePlaceholder'));
      if (!name || !name.trim()) return;
      name = name.trim();
    }
  } else {
    // Overwriting existing template
    name = saveTarget;
  }

  var index = await getTemplateIndex();
  // Update existing or add new
  var found = false;
  for (var i = 0; i < index.templates.length; i++) {
    if (index.templates[i].name === name) {
      index.templates[i].html = templateHtml;
      found = true;
      break;
    }
  }
  if (!found) {
    index.templates.push({ name: name, html: templateHtml });
  }

  await saveTemplateIndex(index);
  currentTemplateName = name;
  if (nameInput) nameInput.value = name;
  await refreshTemplateList();
  showToast(saveTarget === 'new' ? t('templateSaved') : (currentLang === 'fr' ? 'Modèle "' + name + '" mis à jour' : 'Template "' + name + '" updated'), 'success');
}

async function loadSavedTemplate() {
  var index = await getTemplateIndex();

  // Legacy migration: try old per-table templates and migrate to global index
  if (index.templates.length === 0 && selectedTable) {
    var legacyHtml = null;
    try {
      legacyHtml = await grist.widgetApi.getOption('template_' + selectedTable);
    } catch (e) {}
    if (!legacyHtml) {
      try { legacyHtml = localStorage.getItem(TEMPLATE_STORAGE_KEY + selectedTable); } catch (e) {}
    }
    // Also try old per-table index format
    if (!legacyHtml) {
      try {
        var oldIdx = await grist.widgetApi.getOption('templateIndex_' + selectedTable);
        if (oldIdx && Array.isArray(oldIdx.templates) && oldIdx.templates.length > 0) {
          // Migrate all old templates to global index
          index.templates = oldIdx.templates;
          await saveTemplateIndex(index);
        }
      } catch (e) {}
    }
    if (legacyHtml && editorInstance) {
      // Migrate legacy single template
      var legacyName = selectedTable + ' - ' + (currentLang === 'fr' ? 'Modèle importé' : 'Imported template');
      index.templates.push({ name: legacyName, html: legacyHtml });
      await saveTemplateIndex(index);
    }
  }

  if (index.templates.length > 0) {
    // Load the first template by default
    var tpl = index.templates[0];
    setEditorHtml(tpl.html);
    templateHtml = tpl.html;
    currentTemplateName = tpl.name;
    var nameInput = document.getElementById('template-name-input');
    if (nameInput) nameInput.value = tpl.name;
    showToast(t('templateLoaded'), 'info');
  }
  await refreshTemplateList();
}

async function deleteSelectedTemplate() {
  var select = document.getElementById('template-list');
  if (!select || !select.value) return;
  var name = select.value;
  var confirmed = await showModal(t('confirmClearTitle'), t('templateDeleteConfirm').replace('{name}', name));
  if (!confirmed) return;

  var index = await getTemplateIndex();
  index.templates = index.templates.filter(function(tpl) { return tpl.name !== name; });
  await saveTemplateIndex(index);

  currentTemplateName = '';
  var nameInput = document.getElementById('template-name-input');
  if (nameInput) nameInput.value = '';
  editorInstance.value = '';
  templateHtml = '';
  await refreshTemplateList();
  showToast(t('templateDeleted').replace('{name}', name), 'info');
}

async function clearEditor() {
  var confirmed = await showModal(t('confirmClearTitle'), t('confirmClear'));
  if (confirmed && editorInstance) {
    editorInstance.value = '';
    templateHtml = '';
    currentTemplateName = '';
    var nameInput = document.getElementById('template-name-input');
    if (nameInput) nameInput.value = '';
    var select = document.getElementById('template-list');
    if (select) select.value = '';
    var delBtn = document.getElementById('btn-delete-template');
    if (delBtn) delBtn.style.display = 'none';
    
    // Clear saved draft too
    try {
      await grist.widgetApi.setOption('editorDraft', '');
      console.log('Draft cleared');
    } catch (e) {
      console.warn('Could not clear draft:', e);
    }
    
    showToast(t('templateCleared'), 'info');
  }
}

// --- Refresh filters button ---

async function refreshFilters() {
  if (!selectedTable) {
    showToast(currentLang === 'fr' ? 'Sélectionnez d\'abord une table' : 'Select a table first', 'error');
    return;
  }
  
  showToast(currentLang === 'fr' ? 'Actualisation des filtres...' : 'Refreshing filters...', 'info');
  
  // Reload table data
  try {
    var data = await grist.docApi.fetchTable(selectedTable);
    tableData = data;
    
    // Reload views and filters
    await loadViewsForTable(selectedTable);
    
    // Re-render preview
    renderPreview();
    
    showToast(currentLang === 'fr' ? 'Filtres actualisés !' : 'Filters refreshed!', 'success');
  } catch (error) {
    console.error('Error refreshing filters:', error);
    showToast(currentLang === 'fr' ? 'Erreur lors de l\'actualisation' : 'Error refreshing', 'error');
  }
}

// --- PDF template selector ---

async function refreshPdfTemplateList() {
  var select = document.getElementById('pdf-template-select');
  if (!select) return;
  var index = await getTemplateIndex();
  var editorLabel = currentLang === 'fr' ? '-- Modèle actuel de l\'éditeur --' : '-- Current editor template --';
  select.innerHTML = '<option value="">' + editorLabel + '</option>';
  for (var i = 0; i < index.templates.length; i++) {
    var opt = document.createElement('option');
    opt.value = index.templates[i].name;
    opt.textContent = index.templates[i].name;
    select.appendChild(opt);
  }
}

async function onPdfTemplateChange() {
  var select = document.getElementById('pdf-template-select');
  if (!select || !select.value) return;
  // Load selected template into templateHtml for PDF generation
  var index = await getTemplateIndex();
  for (var i = 0; i < index.templates.length; i++) {
    if (index.templates[i].name === select.value) {
      templateHtml = index.templates[i].html;
      return;
    }
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

// =============================================================================
// LOOP PROCESSING - {{#each Column=Value}}...{{/each}}
// =============================================================================

// Sync bool format between preview and PDF selectors
function syncBoolFormat(value) {
  var pdfSelect = document.getElementById('bool-format');
  var previewSelect = document.getElementById('bool-format-preview');
  if (pdfSelect) pdfSelect.value = value;
  if (previewSelect) previewSelect.value = value;
}

function formatValueForDisplay(value) {
  if (value === null || value === undefined || value === '') return '';
  
  // Handle booleans - convert to readable format or checkbox
  // Check both selectors (preview and PDF tab)
  var boolFormatSelect = document.getElementById('bool-format') || document.getElementById('bool-format-preview');
  var useCheckbox = boolFormatSelect && boolFormatSelect.value === 'checkbox';
  
  if (value === true || value === 'true') {
    return useCheckbox ? '☑' : (currentLang === 'fr' ? 'Oui' : 'Yes');
  }
  if (value === false || value === 'false') {
    return useCheckbox ? '☐' : (currentLang === 'fr' ? 'Non' : 'No');
  }
  
  var str = String(value);
  
  // Check if it's a Grist timestamp (number of seconds since epoch, typically 10+ digits)
  if (/^\d{10,}$/.test(str)) {
    var timestamp = parseInt(str);
    var date = new Date(timestamp * 1000);
    if (!isNaN(date.getTime())) {
      var day = String(date.getDate()).padStart(2, '0');
      var month = String(date.getMonth() + 1).padStart(2, '0');
      var year = date.getFullYear();
      return day + '/' + month + '/' + year;
    }
  }
  
  // Check if it's an ISO date format (YYYY-MM-DD) and convert to French format (DD/MM/YYYY)
  var isoDateMatch = str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoDateMatch) {
    var year = isoDateMatch[1];
    var month = isoDateMatch[2];
    var day = isoDateMatch[3];
    return day + '/' + month + '/' + year;
  }
  
  return str;
}

function normalizeForComparison(value) {
  if (!value) return '';
  var str = String(value).trim().toLowerCase();
  
  // Handle Grist timestamp (number of seconds since epoch)
  if (/^\d{10,}$/.test(str)) {
    var date = new Date(parseInt(str) * 1000);
    if (!isNaN(date.getTime())) {
      // Return multiple formats for matching
      var day = String(date.getDate()).padStart(2, '0');
      var month = String(date.getMonth() + 1).padStart(2, '0');
      var year = date.getFullYear();
      var shortYear = String(year).slice(-2);
      return day + '/' + month + '/' + year + '|' + day + '/' + month + '/' + shortYear + '|' + year + '-' + month + '-' + day;
    }
  }
  
  // Handle ISO date format (2026-02-16)
  var isoMatch = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) {
    var y = isoMatch[1], m = isoMatch[2], d = isoMatch[3];
    return d + '/' + m + '/' + y + '|' + d + '/' + m + '/' + y.slice(-2) + '|' + y + '-' + m + '-' + d;
  }
  
  // Handle French date format (16/02/2026 or 16/02/26)
  var frMatch = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (frMatch) {
    var dd = frMatch[1].padStart(2, '0');
    var mm = frMatch[2].padStart(2, '0');
    var yy = frMatch[3];
    if (yy.length === 2) {
      yy = (parseInt(yy) > 50 ? '19' : '20') + yy;
    }
    return dd + '/' + mm + '/' + yy + '|' + dd + '/' + mm + '/' + yy.slice(-2) + '|' + yy + '-' + mm + '-' + dd;
  }
  
  return str;
}

function processLoops(html, forPdf) {
  if (!tableData || !tableColumns.length) return html;
  
  var resolved = html;
  
  // Process HTML comment-based loops (for tables): <!--LOOP:Column=Value-->...<!--/LOOP-->
  // Also supports <!--LOOP:*--> for view-linked tables (all rows)
  // Also supports <!--LOOP:VIEW:viewId--> for specific view filters
  // Also supports <!--LOOP:TABLE:tableName:refColumn--> for linked tables
  var commentLoopRegex = /<!--LOOP:(\*|VIEW:\d+|TABLE:[^:]+:[^-]+|[^=]+=[^-]+)-->([\s\S]*?)<!--\/LOOP-->/gi;
  resolved = resolved.replace(commentLoopRegex, function(match, loopSpec, loopContent) {
    if (loopSpec === '*') {
      // View-linked: show all rows from current view
      return executeLoopAllRows(loopContent, forPdf);
    } else if (loopSpec.startsWith('VIEW:')) {
      // View-select: show rows filtered by specific view
      var viewId = loopSpec.substring(5);
      return executeLoopFromView(viewId, loopContent, forPdf);
    } else if (loopSpec.startsWith('TABLE:')) {
      // Linked table: TABLE:tableName:refColumn
      var tableParts = loopSpec.substring(6).split(':');
      var linkedTableName = tableParts[0].trim();
      var refColumn = tableParts[1].trim();
      return executeLoopLinkedTable(linkedTableName, refColumn, loopContent, forPdf);
    } else {
      // Filtered: parse Column=Value
      var parts = loopSpec.split('=');
      var filterCol = parts[0].trim();
      var filterVal = parts.slice(1).join('=').trim(); // Handle values with = in them
      return executeLoop(filterCol, filterVal, loopContent, forPdf);
    }
  });
  
  // Special case: handle loops inside table rows
  // Pattern: <tr>...<td>{{#each...}}</td>...</tr>...<tr>...</tr>...<tr>...<td>{{/each}}</td>...</tr>
  // This happens when Jodit puts the loop markers in table cells
  var tableLoopRegex = /<tr[^>]*>([^]*?)<td[^>]*>([^<]*\{\{#each\s+([^=}]+)=([^}]+)\}\}[^<]*)<\/td>([^]*?)<\/tr>([^]*?)<tr[^>]*>([^]*?)<\/tr>([^]*?)<tr[^>]*>([^]*?)<td[^>]*>([^<]*\{\{\/each\}\}[^<]*)<\/td>([^]*?)<\/tr>/gi;
  
  resolved = resolved.replace(tableLoopRegex, function(match, before1, eachCell, filterCol, filterVal, after1, between, rowContent, after2, before3, endCell, after3) {
    // Extract the template row (the middle <tr>)
    var templateRow = '<tr>' + rowContent + '</tr>';
    var result = executeLoop(filterCol.trim(), filterVal.trim(), templateRow, forPdf);
    return result;
  });
  
  // Simpler table loop: <tr> containing {{#each}} ... next <tr> with content ... <tr> with {{/each}}
  // Try to detect: row with #each, then row(s) with variables, then row with /each
  var simpleTableLoopRegex = /<tr[^>]*>\s*<td[^>]*>\s*\{\{#each\s+([^=}]+)=([^}]+)\}\}\s*<\/td>\s*<\/tr>\s*(<tr[^>]*>[\s\S]*?<\/tr>)\s*<tr[^>]*>\s*<td[^>]*>\s*\{\{\/each\}\}\s*<\/td>\s*<\/tr>/gi;
  
  resolved = resolved.replace(simpleTableLoopRegex, function(match, filterCol, filterVal, templateRows) {
    return executeLoop(filterCol.trim(), filterVal.trim(), templateRows, forPdf);
  });
  
  // Regex to match {{#each Column=Value}}...{{/each}}
  // Supports both plain text and styled spans
  var loopRegex = /\{\{#each\s+([^=}]+)=([^}]+)\}\}([\s\S]*?)\{\{\/each\}\}/g;
  var styledLoopRegex = /<span[^>]*>\{\{#each\s+([^=}]+)=([^}]+)\}\}<\/span>([\s\S]*?)<span[^>]*>\{\{\/each\}\}<\/span>/g;
  
  // Process styled loops first
  resolved = resolved.replace(styledLoopRegex, function(match, filterCol, filterVal, loopContent) {
    return executeLoop(filterCol.trim(), filterVal.trim(), loopContent, forPdf);
  });
  
  // Process plain text loops
  resolved = resolved.replace(loopRegex, function(match, filterCol, filterVal, loopContent) {
    return executeLoop(filterCol.trim(), filterVal.trim(), loopContent, forPdf);
  });
  
  return resolved;
}

function executeLoopFromView(viewId, loopContent, forPdf) {
  // Execute loop using filters from a specific view
  if (!tableData || !tableColumns.length) return '';
  
  // Find the view in availableViews
  var viewInfo = null;
  for (var i = 0; i < availableViews.length; i++) {
    if (String(availableViews[i].id) === String(viewId)) {
      viewInfo = availableViews[i];
      break;
    }
  }
  
  if (!viewInfo) {
    return '<span style="color:red;">[' + (currentLang === 'fr' ? 'Vue non trouvée: ' : 'View not found: ') + viewId + ']</span>';
  }
  
  // Parse filters from the view
  // viewInfo.filters is now an array of {colRef, filter} objects from _grist_Filters
  var filters = [];
  if (viewInfo.filters && Array.isArray(viewInfo.filters)) {
    try {
      for (var f = 0; f < viewInfo.filters.length; f++) {
        var filterDef = viewInfo.filters[f];
        var colRef = filterDef.colRef;
        var filterJson = filterDef.filter;
        
        // Resolve column name from colRef
        var colName = columnIdToName[colRef];
        if (colName && filterJson) {
          // Parse filter values: {"included":[...]} or {"excluded":[...]}
          var filterData = JSON.parse(filterJson);
          if (filterData.included && filterData.included.length > 0) {
            filters.push({
              column: colName,
              type: 'included',
              values: filterData.included
            });
          } else if (filterData.excluded && filterData.excluded.length > 0) {
            filters.push({
              column: colName,
              type: 'excluded',
              values: filterData.excluded
            });
          }
        }
      }
    } catch (e) {
      console.error('Error parsing view filters:', e);
    }
  }
  
  console.log('Executing loop from view', viewInfo.name, 'with filters:', filters);
  
  // Get the number of rows
  var firstCol = tableColumns[0];
  var rowCount = tableData[firstCol] ? tableData[firstCol].length : 0;
  
  if (rowCount === 0) {
    return '<span style="color:#f59e0b;font-style:italic;">[' + (currentLang === 'fr' ? 'Aucune ligne' : 'No rows') + ']</span>';
  }
  
  // Generate output for each row that matches the filters
  var output = '';
  var matchCount = 0;
  
  for (var j = 0; j < rowCount; j++) {
    var rowRecord = getRecordAt(j);
    
    // Check if row matches all filters
    var matches = true;
    for (var fi = 0; fi < filters.length; fi++) {
      var filter = filters[fi];
      var rowVal = rowRecord[filter.column];
      
      // Check if row value matches the filter
      var found = false;
      for (var vi = 0; vi < filter.values.length; vi++) {
        // Compare as same type (number to number, string to string)
        if (rowVal === filter.values[vi] || String(rowVal) === String(filter.values[vi])) {
          found = true;
          break;
        }
      }
      
      if (filter.type === 'included') {
        // For included filters, row must have one of the values
        if (!found) {
          matches = false;
          break;
        }
      } else if (filter.type === 'excluded') {
        // For excluded filters, row must NOT have any of the values
        if (found) {
          matches = false;
          break;
        }
      }
    }
    
    if (!matches) continue;
    matchCount++;
    
    // Resolve variables in loopContent for this row
    var rowHtml = loopContent;
    for (var col in rowRecord) {
      var val = rowRecord[col];
      var display = formatValueForDisplay(val);
      
      // Replace styled spans
      var styledRegex = new RegExp('<span[^>]*>\\{\\{' + escapeRegex(col) + '\\}\\}</span>', 'g');
      if (display) {
        if (forPdf) {
          rowHtml = rowHtml.replace(styledRegex, sanitize(display));
        } else {
          rowHtml = rowHtml.replace(styledRegex, '<span class="var-resolved">' + sanitize(display) + '</span>');
        }
      } else {
        rowHtml = rowHtml.replace(styledRegex, '');
      }
      
      // Replace plain text variables
      var plainRegex = new RegExp('\\{\\{' + escapeRegex(col) + '\\}\\}', 'g');
      if (display) {
        if (forPdf) {
          rowHtml = rowHtml.replace(plainRegex, sanitize(display));
        } else {
          rowHtml = rowHtml.replace(plainRegex, '<span class="var-resolved">' + sanitize(display) + '</span>');
        }
      } else {
        rowHtml = rowHtml.replace(plainRegex, '');
      }
    }
    output += rowHtml;
  }
  
  if (matchCount === 0) {
    return '<span style="color:#f59e0b;font-style:italic;">[' + (currentLang === 'fr' ? 'Aucune ligne correspondante dans la vue ' : 'No matching rows in view ') + viewInfo.name + ']</span>';
  }
  
  return output;
}

// Cache for linked table data
var linkedTableCache = {};

async function executeLoopLinkedTableAsync(linkedTableName, refColumn, loopContent, forPdf, currentRecordId) {
  // Execute loop for rows from a linked table, filtered by reference to current record
  // Example: Facture_details where Facture_Ref = currentRecordId
  
  var output = '';
  
  try {
    // Fetch linked table data (use cache if available)
    var linkedData;
    if (linkedTableCache[linkedTableName]) {
      linkedData = linkedTableCache[linkedTableName];
    } else {
      linkedData = await grist.docApi.fetchTable(linkedTableName);
      linkedTableCache[linkedTableName] = linkedData;
    }
    
    if (!linkedData || !linkedData.id) {
      return '<span style="color:#ef4444;font-style:italic;">[' + (currentLang === 'fr' ? 'Table non trouvée: ' : 'Table not found: ') + linkedTableName + ']</span>';
    }
    
    // Check if refColumn exists
    if (!linkedData[refColumn]) {
      return '<span style="color:#ef4444;font-style:italic;">[' + (currentLang === 'fr' ? 'Colonne de référence non trouvée: ' : 'Reference column not found: ') + refColumn + ']</span>';
    }
    
    var rowCount = linkedData.id.length;
    var matchCount = 0;
    
    // Get columns from linked table
    var linkedColumns = Object.keys(linkedData).filter(function(k) {
      return k !== 'id' && k !== 'manualSort' && !k.startsWith('gristHelper_');
    });
    
    for (var i = 0; i < rowCount; i++) {
      // Check if this row references the current record
      var refValue = linkedData[refColumn][i];
      
      // Handle both direct ID and reference object
      var refId = refValue;
      if (typeof refValue === 'object' && refValue !== null) {
        refId = refValue.id || refValue;
      }
      
      if (refId !== currentRecordId) continue;
      
      matchCount++;
      
      // Build row record
      var rowRecord = { id: linkedData.id[i] };
      for (var c = 0; c < linkedColumns.length; c++) {
        var col = linkedColumns[c];
        rowRecord[col] = linkedData[col][i];
      }
      
      // Resolve variables in loopContent for this row
      var rowHtml = loopContent;
      for (var col in rowRecord) {
        if (col === 'id') continue;
        var val = rowRecord[col];
        var display = formatValueForDisplay(val);
        
        // Replace styled spans
        var styledRegex = new RegExp('<span[^>]*>\\{\\{' + escapeRegex(col) + '\\}\\}</span>', 'g');
        if (display) {
          if (forPdf) {
            rowHtml = rowHtml.replace(styledRegex, sanitize(display));
          } else {
            rowHtml = rowHtml.replace(styledRegex, '<span class="var-resolved">' + sanitize(display) + '</span>');
          }
        } else {
          rowHtml = rowHtml.replace(styledRegex, '');
        }
        
        // Replace plain text variables
        var plainRegex = new RegExp('\\{\\{' + escapeRegex(col) + '\\}\\}', 'g');
        if (display) {
          if (forPdf) {
            rowHtml = rowHtml.replace(plainRegex, sanitize(display));
          } else {
            rowHtml = rowHtml.replace(plainRegex, '<span class="var-resolved">' + sanitize(display) + '</span>');
          }
        } else {
          rowHtml = rowHtml.replace(plainRegex, '');
        }
      }
      output += rowHtml;
    }
    
    if (matchCount === 0) {
      return '<span style="color:#f59e0b;font-style:italic;">[' + (currentLang === 'fr' ? 'Aucune ligne liée dans ' : 'No linked rows in ') + linkedTableName + ']</span>';
    }
    
    return output;
  } catch (error) {
    console.error('Error in executeLoopLinkedTableAsync:', error);
    return '<span style="color:#ef4444;font-style:italic;">[' + (currentLang === 'fr' ? 'Erreur: ' : 'Error: ') + error.message + ']</span>';
  }
}

function executeLoopLinkedTable(linkedTableName, refColumn, loopContent, forPdf) {
  // Synchronous wrapper - returns placeholder that will be replaced async
  // For PDF, we need to wait for data - this is handled by processLoopsAsync
  // Get current record ID
  if (!tableData || !tableData.id || tableData.id.length === 0) {
    return '<span style="color:#ef4444;font-style:italic;">[' + (currentLang === 'fr' ? 'Pas d\'enregistrement courant' : 'No current record') + ']</span>';
  }
  
  var currentRecordId = tableData.id[currentRecordIndex];
  
  // Create a unique placeholder
  var placeholderId = 'linked-table-' + linkedTableName + '-' + refColumn + '-' + currentRecordId + '-' + Date.now();
  
  // Execute async and update DOM when ready
  executeLoopLinkedTableAsync(linkedTableName, refColumn, loopContent, forPdf, currentRecordId).then(function(result) {
    var placeholder = document.getElementById(placeholderId);
    if (placeholder) {
      // Create a temporary container to parse the result
      var temp = document.createElement('tbody');
      temp.innerHTML = result;
      // Replace placeholder's parent row or the placeholder itself
      var parentTr = placeholder.closest('tr');
      if (parentTr && parentTr.parentNode) {
        // Replace the placeholder row with the new rows
        var parent = parentTr.parentNode;
        while (temp.firstChild) {
          parent.insertBefore(temp.firstChild, parentTr);
        }
        parent.removeChild(parentTr);
      } else {
        // Fallback: just replace the span
        placeholder.outerHTML = result;
      }
    }
  });
  
  // Return a placeholder row instead of a span
  return '<tr id="' + placeholderId + '"><td colspan="99" style="color:#6b7280;font-style:italic;text-align:center;">' + (currentLang === 'fr' ? 'Chargement...' : 'Loading...') + '</td></tr>';
}

// Async version of processLoops for PDF generation - waits for linked table data
async function processLoopsAsync(html, forPdf) {
  if (!tableData || !tableColumns.length) return html;
  
  var resolved = html;
  
  // Process HTML comment-based loops (for tables): <!--LOOP:Column=Value-->...<!--/LOOP-->
  // Handle TABLE: loops specially - they need async processing
  var commentLoopRegex = /<!--LOOP:(\*|VIEW:\d+|TABLE:[^:]+:[^-]+|[^=]+=[^-]+)-->([\s\S]*?)<!--\/LOOP-->/gi;
  
  // Collect all matches first
  var matches = [];
  var match;
  var tempHtml = resolved;
  var regex = new RegExp(commentLoopRegex.source, 'gi');
  while ((match = regex.exec(tempHtml)) !== null) {
    matches.push({
      fullMatch: match[0],
      loopSpec: match[1],
      loopContent: match[2],
      index: match.index
    });
  }
  
  // Process matches in reverse order to preserve indices
  for (var i = matches.length - 1; i >= 0; i--) {
    var m = matches[i];
    var replacement;
    
    if (m.loopSpec === '*') {
      replacement = executeLoopAllRows(m.loopContent, forPdf);
    } else if (m.loopSpec.startsWith('VIEW:')) {
      var viewId = m.loopSpec.substring(5);
      replacement = executeLoopFromView(viewId, m.loopContent, forPdf);
    } else if (m.loopSpec.startsWith('TABLE:')) {
      // Linked table: TABLE:tableName:refColumn - AWAIT the async function
      var tableParts = m.loopSpec.substring(6).split(':');
      var linkedTableName = tableParts[0].trim();
      var refColumn = tableParts[1].trim();
      
      // Get current record ID
      var currentRecordId = tableData.id[currentRecordIndex];
      // Await the async function directly for PDF
      replacement = await executeLoopLinkedTableAsync(linkedTableName, refColumn, m.loopContent, forPdf, currentRecordId);
    } else {
      // Filtered: parse Column=Value
      var parts = m.loopSpec.split('=');
      var filterCol = parts[0].trim();
      var filterVal = parts.slice(1).join('=').trim();
      replacement = executeLoop(filterCol, filterVal, m.loopContent, forPdf);
    }
    
    resolved = resolved.substring(0, m.index) + replacement + resolved.substring(m.index + m.fullMatch.length);
  }
  
  // Process other loop types (same as sync version)
  // Special case: handle loops inside table rows
  var tableLoopRegex = /<tr[^>]*>([^]*?)<td[^>]*>([^<]*\{\{#each\s+([^=}]+)=([^}]+)\}\}[^<]*)<\/td>([^]*?)<\/tr>([^]*?)<tr[^>]*>([^]*?)<\/tr>([^]*?)<tr[^>]*>([^]*?)<td[^>]*>([^<]*\{\{\/each\}\}[^<]*)<\/td>([^]*?)<\/tr>/gi;
  
  resolved = resolved.replace(tableLoopRegex, function(match, before1, eachCell, filterCol, filterVal, after1, between, rowContent, after2, before3, endCell, after3) {
    var templateRow = '<tr>' + rowContent + '</tr>';
    var result = executeLoop(filterCol.trim(), filterVal.trim(), templateRow, forPdf);
    return result;
  });
  
  // Simpler table loop
  var simpleTableLoopRegex = /<tr[^>]*>\s*<td[^>]*>\s*\{\{#each\s+([^=}]+)=([^}]+)\}\}\s*<\/td>\s*<\/tr>\s*(<tr[^>]*>[\s\S]*?<\/tr>)\s*<tr[^>]*>\s*<td[^>]*>\s*\{\{\/each\}\}\s*<\/td>\s*<\/tr>/gi;
  
  resolved = resolved.replace(simpleTableLoopRegex, function(match, filterCol, filterVal, templateRows) {
    return executeLoop(filterCol.trim(), filterVal.trim(), templateRows, forPdf);
  });
  
  // Regex to match {{#each Column=Value}}...{{/each}}
  var loopRegex = /\{\{#each\s+([^=}]+)=([^}]+)\}\}([\s\S]*?)\{\{\/each\}\}/g;
  var styledLoopRegex = /<span[^>]*>\{\{#each\s+([^=}]+)=([^}]+)\}\}<\/span>([\s\S]*?)<span[^>]*>\{\{\/each\}\}<\/span>/g;
  
  resolved = resolved.replace(styledLoopRegex, function(match, filterCol, filterVal, loopContent) {
    return executeLoop(filterCol.trim(), filterVal.trim(), loopContent, forPdf);
  });
  
  resolved = resolved.replace(loopRegex, function(match, filterCol, filterVal, loopContent) {
    return executeLoop(filterCol.trim(), filterVal.trim(), loopContent, forPdf);
  });
  
  return resolved;
}

// Async version of resolveTemplate for PDF generation
async function resolveTemplateAsync(html, record, forPdf) {
  var resolved = html;
  
  // Process loops with async support for linked tables
  resolved = await processLoopsAsync(resolved, forPdf);
  
  for (var col in record) {
    var val = record[col];
    var display = formatValueForDisplay(val);
    var styledRegex = new RegExp('<span[^>]*>\\{\\{' + escapeRegex(col) + '\\}\\}</span>', 'g');
    if (display) {
      if (forPdf) {
        resolved = resolved.replace(styledRegex, sanitize(display));
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
    var plainRegex = new RegExp('\\{\\{' + escapeRegex(col) + '\\}\\}', 'g');
    if (display) {
      if (forPdf) {
        resolved = resolved.replace(plainRegex, sanitize(display));
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
  
  if (forPdf) {
    resolved = resolved.replace(/<div[^>]*class="page-break-marker"[^>]*>[\s\S]*?<\/div>/g, '<div style="page-break-after:always;"></div>');
    resolved = resolved.replace(/background-color:\s*rgb\(243,\s*232,\s*255\);?/g, '');
    resolved = resolved.replace(/background-color:\s*#f3e8ff;?/g, '');
    resolved = resolved.replace(/color:\s*rgb\(124,\s*58,\s*237\);?/g, '');
    resolved = resolved.replace(/color:\s*#7c3aed;?/g, '');
    resolved = resolved.replace(/<table(?![^>]*style=)/g, '<table style="border-collapse:collapse;"');
  }
  return resolved;
}

function executeLoopAllRows(loopContent, forPdf) {
  // Execute loop for rows from "Select By" filtered view
  // If no filtered records, fall back to all tableData rows
  
  var recordsToUse = filteredRecords.length > 0 ? filteredRecords : null;
  
  // If we have filtered records from "Select By", use them
  if (recordsToUse && recordsToUse.length > 0) {
    console.log('Using filtered records from Select By:', recordsToUse.length);
    var output = '';
    for (var j = 0; j < recordsToUse.length; j++) {
      var rowRecord = recordsToUse[j];
      
      // Resolve variables in loopContent for this row
      var rowHtml = loopContent;
      for (var col in rowRecord) {
        if (col === 'id') continue;
        var val = rowRecord[col];
        var display = formatValueForDisplay(val);
        
        // Replace styled spans
        var styledRegex = new RegExp('<span[^>]*>\\{\\{' + escapeRegex(col) + '\\}\\}</span>', 'g');
        if (display) {
          if (forPdf) {
            rowHtml = rowHtml.replace(styledRegex, sanitize(display));
          } else {
            rowHtml = rowHtml.replace(styledRegex, '<span class="var-resolved">' + sanitize(display) + '</span>');
          }
        } else {
          rowHtml = rowHtml.replace(styledRegex, '');
        }
        
        // Replace plain text variables
        var plainRegex = new RegExp('\\{\\{' + escapeRegex(col) + '\\}\\}', 'g');
        if (display) {
          if (forPdf) {
            rowHtml = rowHtml.replace(plainRegex, sanitize(display));
          } else {
            rowHtml = rowHtml.replace(plainRegex, '<span class="var-resolved">' + sanitize(display) + '</span>');
          }
        } else {
          rowHtml = rowHtml.replace(plainRegex, '');
        }
      }
      output += rowHtml;
    }
    return output;
  }
  
  // Fallback: use all rows from tableData
  if (!tableData || !tableColumns.length) return '';
  
  // Get the number of rows from the first column
  var firstCol = tableColumns[0];
  var rowCount = tableData[firstCol] ? tableData[firstCol].length : 0;
  
  if (rowCount === 0) {
    return '<span style="color:#f59e0b;font-style:italic;">[' + (currentLang === 'fr' ? 'Aucune ligne dans la vue' : 'No rows in view') + ']</span>';
  }
  
  console.log('Using all rows from tableData (no Select By filter):', rowCount);
  
  // Generate output for each row
  var output = '';
  for (var j = 0; j < rowCount; j++) {
    var rowRecord = getRecordAt(j);
    
    // Resolve variables in loopContent for this row
    var rowHtml = loopContent;
    for (var col in rowRecord) {
      var val = rowRecord[col];
      var display = formatValueForDisplay(val);
      
      // Replace styled spans
      var styledRegex = new RegExp('<span[^>]*>\\{\\{' + escapeRegex(col) + '\\}\\}</span>', 'g');
      if (display) {
        if (forPdf) {
          rowHtml = rowHtml.replace(styledRegex, sanitize(display));
        } else {
          rowHtml = rowHtml.replace(styledRegex, '<span class="var-resolved">' + sanitize(display) + '</span>');
        }
      } else {
        rowHtml = rowHtml.replace(styledRegex, '');
      }
      
      // Replace plain text variables
      var plainRegex = new RegExp('\\{\\{' + escapeRegex(col) + '\\}\\}', 'g');
      if (display) {
        if (forPdf) {
          rowHtml = rowHtml.replace(plainRegex, sanitize(display));
        } else {
          rowHtml = rowHtml.replace(plainRegex, '<span class="var-resolved">' + sanitize(display) + '</span>');
        }
      } else {
        rowHtml = rowHtml.replace(plainRegex, '');
      }
    }
    output += rowHtml;
  }
  
  return output;
}

function executeLoop(filterColumn, filterValue, loopContent, forPdf) {
  if (!tableData || !tableColumns.length) return '';
  
  // Find the filter column in tableData
  var filterColData = tableData[filterColumn];
  if (!filterColData) {
    // Column not found - return error message with available columns
    var availableCols = tableColumns.join(', ');
    return '<span style="color:red;">[Colonne "' + filterColumn + '" non trouvée. Colonnes disponibles: ' + availableCols + ']</span>';
  }
  
  // Find all rows where filterColumn matches filterValue
  var matchingIndices = [];
  var sampleValues = [];
  
  for (var i = 0; i < filterColData.length; i++) {
    var cellValue = filterColData[i];
    var cellStr = (cellValue === null || cellValue === undefined) ? '' : String(cellValue);
    
    // Collect sample values for debug (first 3 unique)
    if (sampleValues.length < 3 && sampleValues.indexOf(cellStr) === -1) {
      sampleValues.push(cellStr);
    }
    
    // Normalize for date comparison
    var normalizedCell = normalizeForComparison(cellStr);
    var normalizedFilter = normalizeForComparison(filterValue);
    
    // Check if filter value is a reference display value (e.g., "DUMZ 60")
    // and if so, check if the resolved value matches
    var refMatch = false;
    var meta = columnMetadata[filterColumn];
    if (meta && meta.type) {
      var refTypeMatch = meta.type.match(/^Ref:(.+)$/);
      if (refTypeMatch) {
        var refTableName = refTypeMatch[1];
        var refDisplayData = referenceDisplayValues[refTableName];
        if (refDisplayData) {
          // Check both byVisibleCol and byFirstTextCol for matching filter value
          var allRefMaps = [refDisplayData.byVisibleCol, refDisplayData.byFirstTextCol];
          for (var mapIdx = 0; mapIdx < allRefMaps.length && !refMatch; mapIdx++) {
            var refMap = allRefMaps[mapIdx];
            if (!refMap) continue;
            for (var refId in refMap) {
              var refDisplayVal = refMap[refId];
              if (refDisplayVal === filterValue || 
                  normalizeForComparison(refDisplayVal) === normalizedFilter) {
                // Check if the resolved cell value matches this reference's resolved value
                var resolvedRefVal = lookupRefValue(referenceTables[refTableName], parseInt(refId), findDisplayColumn(referenceTables[refTableName], meta.visibleCol));
                if (cellStr === resolvedRefVal || normalizedCell === normalizeForComparison(resolvedRefVal)) {
                  refMatch = true;
                  break;
                }
              }
            }
          }
        }
      }
    }
    
    // Flexible matching: exact match, contains, normalized match, or reference match
    if (cellStr === filterValue || 
        cellStr.indexOf(filterValue) !== -1 ||
        normalizedCell === normalizedFilter ||
        normalizedCell.indexOf(normalizedFilter) !== -1 ||
        refMatch) {
      matchingIndices.push(i);
    }
  }
  
  if (matchingIndices.length === 0) {
    // No matches found - show sample values to help user
    var sampleStr = sampleValues.map(function(v) { return '"' + v + '"'; }).join(', ');
    return '<span style="color:#f59e0b;font-style:italic;">[Aucune ligne où ' + filterColumn + '="' + filterValue + '". Valeurs existantes: ' + sampleStr + '...]</span>';
  }
  
  // Generate output for each matching row
  var output = '';
  for (var j = 0; j < matchingIndices.length; j++) {
    var rowIndex = matchingIndices[j];
    var rowRecord = getRecordAt(rowIndex);
    
    // Resolve variables in loopContent for this row
    var rowHtml = loopContent;
    for (var col in rowRecord) {
      var val = rowRecord[col];
      var display = formatValueForDisplay(val);
      
      // Replace styled spans
      var styledRegex = new RegExp('<span[^>]*>\\{\\{' + escapeRegex(col) + '\\}\\}</span>', 'g');
      if (display) {
        if (forPdf) {
          rowHtml = rowHtml.replace(styledRegex, sanitize(display));
        } else {
          rowHtml = rowHtml.replace(styledRegex, '<span class="var-resolved">' + sanitize(display) + '</span>');
        }
      } else {
        rowHtml = rowHtml.replace(styledRegex, '');
      }
      
      // Replace plain text {{col}}
      var plainRegex = new RegExp('\\{\\{' + escapeRegex(col) + '\\}\\}', 'g');
      if (display) {
        if (forPdf) {
          rowHtml = rowHtml.replace(plainRegex, sanitize(display));
        } else {
          rowHtml = rowHtml.replace(plainRegex, '<span class="var-resolved">' + sanitize(display) + '</span>');
        }
      } else {
        rowHtml = rowHtml.replace(plainRegex, '');
      }
    }
    
    output += rowHtml;
  }
  
  return output;
}

function resolveTemplate(html, record, forPdf) {
  var resolved = html;
  
  // Process {{#each Column=Value}}...{{/each}} loops first
  resolved = processLoops(resolved, forPdf);
  
  for (var col in record) {
    var val = record[col];
    var display = formatValueForDisplay(val);
    // Replace styled spans (Quill wraps variables in <span style="...">)
    var styledRegex = new RegExp('<span[^>]*>\\{\\{' + escapeRegex(col) + '\\}\\}</span>', 'g');
    if (display) {
      if (forPdf) {
        resolved = resolved.replace(styledRegex, sanitize(display));
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
        resolved = resolved.replace(plainRegex, sanitize(display));
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
  // For PDF: remove page-break markers (keep them only as split points, not visible)
  if (forPdf) {
    resolved = resolved.replace(/<div[^>]*class="page-break-marker"[^>]*>[\s\S]*?<\/div>/g, '<div style="page-break-after:always;"></div>');
  }
  // For PDF: strip only variable styling (purple colors), keep table styles
  if (forPdf) {
    // Remove variable highlight colors only
    resolved = resolved.replace(/background-color:\s*rgb\(243,\s*232,\s*255\);?/g, '');
    resolved = resolved.replace(/background-color:\s*#f3e8ff;?/g, '');
    resolved = resolved.replace(/color:\s*rgb\(124,\s*58,\s*237\);?/g, '');
    resolved = resolved.replace(/color:\s*#7c3aed;?/g, '');
    // Keep table styles intact - only add border-collapse if missing
    resolved = resolved.replace(/<table(?![^>]*style=)/g, '<table style="border-collapse:collapse;"');
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

  var wrapper = document.getElementById('preview-pages-wrapper');

  // Sync preview format with editor format
  var editorFormat = document.getElementById('editor-page-format');
  if (editorFormat && wrapper) {
    var fmt = editorFormat.value;
    wrapper.className = 'preview-pages-wrapper';
    if (fmt === 'a4') wrapper.classList.add('preview-format-a4');
    else if (fmt === 'letter') wrapper.classList.add('preview-format-letter');
  }

  if (!templateHtml || !selectedTable || count === 0) {
    wrapper.innerHTML = '<div class="preview-page"><p style="color:#94a3b8; text-align:center; padding:40px;">' + t('previewEmpty') + '</p></div>';
    document.getElementById('record-current').textContent = '0';
    return;
  }

  if (currentRecordIndex >= count) currentRecordIndex = count - 1;
  if (currentRecordIndex < 0) currentRecordIndex = 0;

  document.getElementById('record-current').textContent = currentRecordIndex + 1;

  var record = getRecordAt(currentRecordIndex);
  var resolved = resolveTemplate(templateHtml, record);

  // Split at page break markers to show separate visual pages
  var pages = splitPreviewIntoPages(resolved);
  var html = '';
  for (var p = 0; p < pages.length; p++) {
    html += '<div class="preview-page">' + pages[p] + '</div>';
    if (p < pages.length - 1) {
      html += '<div class="preview-page-number">' +
        (currentLang === 'fr' ? 'Page ' : 'Page ') + (p + 1) + ' / ' + pages.length +
        '</div>';
    }
  }
  if (pages.length > 1) {
    html += '<div class="preview-page-number">' +
      (currentLang === 'fr' ? 'Page ' : 'Page ') + pages.length + ' / ' + pages.length +
      '</div>';
  }
  wrapper.innerHTML = html;
}

function splitPreviewIntoPages(html) {
  // First split on explicit page-break markers
  var parts = html.split(/<div[^>]*class="page-break-marker"[^>]*>[\s\S]*?<\/div>/g);
  parts = parts.reduce(function(acc, part) {
    var subParts = part.split(/<div[^>]*style="[^"]*page-break-after:\s*always[^"]*"[^>]*>\s*<\/div>/g);
    return acc.concat(subParts);
  }, []);
  parts = parts.reduce(function(acc, part) {
    var subParts = part.split(/<hr[^>]*style="[^"]*page-break[^"]*"[^>]*\/?>/g);
    return acc.concat(subParts);
  }, []);
  parts = parts.filter(function(p) { return p.trim().length > 0; });

  // Preview shows content as-is (no automatic height-based pagination)
  // The PDF generation handles pagination correctly
  // Preview only splits on explicit page-break markers
  return parts;
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
  // Use async version to wait for linked table data
  var resolved = await resolveTemplateAsync(templateHtml, record, true);
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
      // Use async version to wait for linked table data
      var resolved = await resolveTemplateAsync(templateHtml, record, true);

      // Render using block-aware page breaking
      if (i > startIdx) {
        pdf.addPage();
      }
      await renderHtmlToPdfPages(resolved, pdf, pageWidth, pageHeight, pageSize);

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

    await renderHtmlToPdfPages(html, pdf, pageWidth, pageHeight, pageSize);

    pdf.save(filename);
    showToast('PDF généré !', 'success');

  } catch (error) {
    console.error('PDF error:', error);
    showToast(t('pdfError') + error.message, 'error');
  }
}

// =============================================================================
// BLOCK-AWARE PDF PAGE RENDERING
// =============================================================================

async function renderHtmlToPdfPages(html, pdf, pageWidth, pageHeight, pageSize) {
  var margin = 10; // mm
  var imgWidth = pageWidth - (margin * 2);
  var availableHeight = pageHeight - (margin * 2);
  var pixelWidth = (pageSize === 'a4' ? 794 : 816);
  var baseCss = 'position:absolute;left:-9999px;top:0;width:' + pixelWidth + 'px;padding:0 60px;font-family:"Times New Roman",Times,serif;font-size:14px;line-height:1.6;background:white;';

  // Split HTML into blocks at page-break markers and top-level elements
  var sections = splitHtmlIntoPageSections(html);
  var currentY = margin;
  var isFirstOnPage = true;

  for (var s = 0; s < sections.length; s++) {
    var section = sections[s];

    // Handle explicit page break
    if (section.isPageBreak) {
      pdf.addPage();
      currentY = margin;
      isFirstOnPage = true;
      continue;
    }

    // Render this block to canvas
    var tempDiv = document.createElement('div');
    tempDiv.style.cssText = baseCss;
    tempDiv.innerHTML = section.html;
    
    // Pre-style tables and cells before appending to DOM
    var tables = tempDiv.querySelectorAll('table');
    tables.forEach(function(table) {
      table.style.width = '100%';
      table.style.borderCollapse = 'collapse';
      table.style.tableLayout = 'auto';
      table.style.pageBreakInside = 'avoid';
    });
    var cells = tempDiv.querySelectorAll('td, th');
    cells.forEach(function(cell) {
      if (!cell.style.border || cell.style.border === 'none') {
        cell.style.border = '1px solid #000';
      }
      if (!cell.style.padding) {
        cell.style.padding = '4px 8px';
      }
    });
    
    // Pre-style images
    var images = tempDiv.querySelectorAll('img');
    images.forEach(function(img) {
      img.style.maxWidth = '100%';
      img.style.height = 'auto';
    });
    
    document.body.appendChild(tempDiv);

    // Convert images to data URLs to avoid CORS issues
    if (images.length > 0) {
      await Promise.all(Array.from(images).map(async function(img) {
        var originalSrc = img.src;
        if (!originalSrc || originalSrc.indexOf('data:') === 0) {
          return; // Already a data URL or no src
        }
        
        try {
          // Try to fetch the image and convert to data URL
          var response = await fetch(originalSrc, { mode: 'cors', credentials: 'include' });
          if (response.ok) {
            var blob = await response.blob();
            var dataUrl = await new Promise(function(resolve) {
              var reader = new FileReader();
              reader.onloadend = function() { resolve(reader.result); };
              reader.readAsDataURL(blob);
            });
            img.src = dataUrl;
          }
        } catch (e) {
          console.warn('Could not convert image to data URL:', originalSrc, e);
          // Try loading normally as fallback
        }
        
        // Wait for image to be ready
        if (!img.complete) {
          await new Promise(function(resolve) {
            img.onload = resolve;
            img.onerror = function() {
              console.warn('Image failed to load:', originalSrc);
              resolve();
            };
          });
        }
      }));
    }

    // Give extra time for tables and complex layouts to render
    await new Promise(function(resolve) { setTimeout(resolve, 150); });

    var canvas = await html2canvas(tempDiv, {
      scale: 2,
      useCORS: true,
      allowTaint: true,
      logging: false,
      backgroundColor: '#ffffff',
      scrollX: 0,
      scrollY: 0,
      windowWidth: tempDiv.scrollWidth,
      windowHeight: tempDiv.scrollHeight,
      onclone: function(clonedDoc, clonedElement) {
        // Ensure tables are fully visible in the cloned element
        var clonedTables = clonedElement.querySelectorAll('table');
        clonedTables.forEach(function(table) {
          table.style.width = '100%';
          table.style.borderCollapse = 'collapse';
          table.style.tableLayout = 'auto';
          table.style.display = 'table';
          table.style.visibility = 'visible';
        });
        // Ensure all rows are visible
        var clonedRows = clonedElement.querySelectorAll('tr');
        clonedRows.forEach(function(row) {
          row.style.display = 'table-row';
          row.style.visibility = 'visible';
        });
        // Ensure all cells have proper borders and are visible
        var clonedCells = clonedElement.querySelectorAll('td, th');
        clonedCells.forEach(function(cell) {
          cell.style.border = '1px solid #000';
          cell.style.padding = '4px 8px';
          cell.style.display = 'table-cell';
          cell.style.visibility = 'visible';
        });
        // Ensure images are visible
        var clonedImages = clonedElement.querySelectorAll('img');
        clonedImages.forEach(function(img) {
          img.style.maxWidth = '100%';
          img.style.visibility = 'visible';
        });
      }
    });

    document.body.removeChild(tempDiv);

    var blockImgHeight = (canvas.height * imgWidth) / canvas.width;
    var imgData = canvas.toDataURL('image/jpeg', 0.95);

    // If block fits on current page - render at full size
    if (currentY + blockImgHeight <= pageHeight - margin) {
      pdf.addImage(imgData, 'JPEG', margin, currentY, imgWidth, blockImgHeight);
      currentY += blockImgHeight;
      isFirstOnPage = false;
    }
    // Block doesn't fit on current page - go to next page
    else {
      if (!isFirstOnPage) {
        pdf.addPage();
        currentY = margin;
      }
      // Render at full width, let it overflow if needed (will be cropped by PDF)
      pdf.addImage(imgData, 'JPEG', margin, currentY, imgWidth, blockImgHeight);
      currentY += blockImgHeight;
      isFirstOnPage = false;
    }
  }
}

function splitHtmlIntoPageSections(html) {
  // Create a temporary container to parse the HTML
  var container = document.createElement('div');
  container.innerHTML = html;

  var sections = [];
  var currentHtml = '';
  var hasExplicitPageBreak = false;

  var children = container.childNodes;
  for (var i = 0; i < children.length; i++) {
    var node = children[i];

    // Check for explicit page break (hr with page-break-after or page-break-before)
    if (node.nodeType === 1) { // Element node
      var style = node.getAttribute('style') || '';
      var tagName = node.tagName.toLowerCase();

      // Detect page break elements (manual markers, Word imports, hr with page-break)
      var classList = node.className || '';
      if (classList.indexOf('page-break-marker') !== -1 ||
          style.indexOf('page-break-after') !== -1 ||
          style.indexOf('page-break-before') !== -1 ||
          (tagName === 'hr' && style.indexOf('page-break') !== -1)) {
        // Push current accumulated content
        if (currentHtml.trim()) {
          sections.push({ html: currentHtml, isPageBreak: false });
          currentHtml = '';
        }
        sections.push({ html: '', isPageBreak: true });
        hasExplicitPageBreak = true;
        continue;
      }
    }

    // Accumulate ALL content (including tables) as one block
    // This prevents unnecessary page breaks
    if (node.nodeType === 1) {
      currentHtml += node.outerHTML;
    } else if (node.nodeType === 3 && node.textContent.trim()) {
      currentHtml += node.textContent;
    }
  }

  // Push remaining content
  if (currentHtml.trim()) {
    sections.push({ html: currentHtml, isPageBreak: false });
  }

  // If no sections were created, return the whole HTML as one section
  if (sections.length === 0) {
    sections.push({ html: html, isPageBreak: false });
  }

  return sections;
}
