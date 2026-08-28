const rootElement = document.getElementById('root');

const state = {
  categoryImport: null,
  permissionImport: null,
  model: null,
  issues: [],
  loadingMessage: '',
  mode: 'categories',
  selectedCategoryId: null,
  selectedPartnerId: null,
  expanded: new Set(),
  filter: 'all',
  search: ''
};

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function makeIssue(severity, source, message, details = {}) {
  return { severity, source, message, ...details };
}

function readField(record, names) {
  const keys = Object.keys(record);
  for (const name of names) {
    const key = keys.find(candidate => candidate.toLowerCase() === name.toLowerCase());
    if (key !== undefined) return { found: true, value: record[key] };
  }
  return { found: false, value: undefined };
}

function parseInteger(value) {
  if (typeof value === 'number' && Number.isSafeInteger(value)) return value;
  if (typeof value === 'string' && /^-?\d+$/.test(value.trim())) {
    const parsed = Number(value.trim());
    if (Number.isSafeInteger(parsed)) return parsed;
  }
  return null;
}

async function parseCategoryFile(file) {
  const issues = [];
  let parsed;

  try {
    parsed = JSON.parse(await file.text());
  } catch (error) {
    return {
      fileName: file.name,
      categories: [],
      issues: [makeIssue('error', 'JSON', `Не удалось прочитать JSON: ${error.message}`)],
      structurallyValid: false
    };
  }

  const records = Array.isArray(parsed) ? parsed : parsed && Array.isArray(parsed.categories) ? parsed.categories : null;
  if (!records) {
    return {
      fileName: file.name,
      categories: [],
      issues: [makeIssue('error', 'JSON', 'Ожидался массив категорий или объект с полем categories.')],
      structurallyValid: false
    };
  }

  const categories = [];
  const seenIds = new Set();

  records.forEach((record, index) => {
    const position = index + 1;
    if (!record || typeof record !== 'object' || Array.isArray(record)) {
      issues.push(makeIssue('error', 'JSON', 'Категория должна быть объектом.', { row: position }));
      return;
    }

    const idField = readField(record, ['categoryId']);
    const parentField = readField(record, ['parentid', 'parentId']);
    const nameField = readField(record, ['name']);
    const displayNameField = readField(record, ['displayname', 'displayName']);
    const treePathField = readField(record, ['treepath', 'treePath']);
    const categoryId = parseInteger(idField.value);

    if (!idField.found || categoryId === null) {
      issues.push(makeIssue('error', 'JSON', 'categoryId должен быть целым числом.', { row: position }));
      return;
    }

    if (seenIds.has(categoryId)) {
      issues.push(makeIssue('error', 'JSON', `Повторяющийся categoryId ${categoryId}.`, { row: position, categoryId }));
      return;
    }
    seenIds.add(categoryId);

    let parentId = null;
    if (parentField.found && parentField.value !== null && parentField.value !== '') {
      parentId = parseInteger(parentField.value);
      if (parentId === null) {
        issues.push(makeIssue('error', 'JSON', 'parentid должен быть целым числом или null.', { row: position, categoryId }));
        return;
      }
    }

    let treePath = [];
    if (!treePathField.found || !Array.isArray(treePathField.value)) {
      issues.push(makeIssue('warning', 'JSON', 'treepath отсутствует или не является массивом.', { row: position, categoryId }));
    } else {
      const parsedPath = treePathField.value.map(parseInteger);
      if (parsedPath.some(value => value === null)) {
        issues.push(makeIssue('warning', 'JSON', 'treepath содержит нецелочисленное значение.', { row: position, categoryId }));
      } else {
        treePath = parsedPath;
      }
    }

    const name = nameField.found && nameField.value != null ? String(nameField.value).trim() : '';
    const displayName = displayNameField.found && displayNameField.value != null ? String(displayNameField.value).trim() : '';
    if (!displayName && !name) {
      issues.push(makeIssue('warning', 'JSON', 'У категории нет displayname и name; будет показан ID.', { row: position, categoryId }));
    }

    categories.push({ categoryId, parentId, name, displayName, treePath, sourceRow: position });
  });

  const categoryMap = new Map(categories.map(category => [category.categoryId, category]));
  categories.forEach(category => {
    if (category.parentId !== null && !categoryMap.has(category.parentId)) {
      issues.push(makeIssue('error', 'JSON', `Родитель ${category.parentId} отсутствует в дереве.`, {
        row: category.sourceRow,
        categoryId: category.categoryId
      }));
    }
  });

  const visitState = new Map();
  function visit(category) {
    const currentState = visitState.get(category.categoryId) || 0;
    if (currentState === 1) {
      issues.push(makeIssue('error', 'JSON', 'Обнаружен цикл в parentid.', {
        row: category.sourceRow,
        categoryId: category.categoryId
      }));
      return;
    }
    if (currentState === 2) return;
    visitState.set(category.categoryId, 1);
    if (category.parentId !== null && categoryMap.has(category.parentId)) visit(categoryMap.get(category.parentId));
    visitState.set(category.categoryId, 2);
  }
  categories.forEach(visit);

  const hasStructuralErrors = issues.some(issue => issue.severity === 'error');
  if (!hasStructuralErrors) {
    categories.forEach(category => {
      const actualPath = [];
      let parentId = category.parentId;
      while (parentId !== null) {
        actualPath.unshift(parentId);
        parentId = categoryMap.get(parentId).parentId;
      }
      if (JSON.stringify(actualPath) !== JSON.stringify(category.treePath)) {
        issues.push(makeIssue('warning', 'JSON', `treepath не совпадает с цепочкой parentid: ожидалось [${actualPath.join(', ')}].`, {
          row: category.sourceRow,
          categoryId: category.categoryId
        }));
      }
    });
  }

  return {
    fileName: file.name,
    categories,
    issues,
    structurallyValid: !issues.some(issue => issue.severity === 'error')
  };
}

async function parsePermissionFile(file) {
  const issues = [];
  let workbook;

  try {
    workbook = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: false });
  } catch (error) {
    return {
      fileName: file.name,
      sheetName: '',
      instructions: [],
      issues: [makeIssue('error', 'Excel', `Не удалось прочитать Excel: ${error.message}`)]
    };
  }

  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    return {
      fileName: file.name,
      sheetName: '',
      instructions: [],
      issues: [makeIssue('error', 'Excel', 'В книге нет листов.')]
    };
  }

  const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
    header: 1,
    defval: null,
    raw: true,
    blankrows: true
  });

  const headerIndex = matrix.findIndex(row => Array.isArray(row) && row.some(value => value !== null && String(value).trim() !== ''));
  if (headerIndex < 0) {
    return {
      fileName: file.name,
      sheetName,
      instructions: [],
      issues: [makeIssue('error', 'Excel', 'Первый лист пуст.')]
    };
  }

  const headers = matrix[headerIndex].map(value => String(value ?? '').trim().toLowerCase());
  const partnerIndex = headers.indexOf('partnerid');
  const categoryIndex = headers.includes('categoryids') ? headers.indexOf('categoryids') : headers.indexOf('categoryid');
  const allowIndex = headers.indexOf('allow');

  if (partnerIndex < 0 || categoryIndex < 0 || allowIndex < 0) {
    const missing = [];
    if (partnerIndex < 0) missing.push('partnerId');
    if (categoryIndex < 0) missing.push('categoryIds');
    if (allowIndex < 0) missing.push('allow');
    return {
      fileName: file.name,
      sheetName,
      instructions: [],
      issues: [makeIssue('error', 'Excel', `Не найдены обязательные столбцы: ${missing.join(', ')}.`, { row: headerIndex + 1 })]
    };
  }

  const instructions = [];
  for (let index = headerIndex + 1; index < matrix.length; index++) {
    const row = matrix[index] || [];
    const sourceRow = index + 1;
    if (row.every(value => value === null || String(value).trim() === '')) continue;

    const partnerId = row[partnerIndex] == null ? '' : String(row[partnerIndex]).trim();
    const categoryId = parseInteger(row[categoryIndex]);
    const allow = row[allowIndex] == null ? '' : String(row[allowIndex]).trim().toLowerCase();
    let valid = true;

    if (!partnerId) {
      issues.push(makeIssue('error', 'Excel', 'partnerId не заполнен.', { row: sourceRow }));
      valid = false;
    }
    if (categoryId === null) {
      issues.push(makeIssue('error', 'Excel', 'categoryIds должен содержать один целочисленный categoryId.', {
        row: sourceRow,
        partnerId
      }));
      valid = false;
    }
    if (allow !== 'yes' && allow !== 'no') {
      issues.push(makeIssue('error', 'Excel', 'allow должен быть yes или no.', {
        row: sourceRow,
        partnerId,
        categoryId
      }));
      valid = false;
    }
    if (valid) instructions.push({ partnerId, categoryId, allow, sourceRow });
  }

  return { fileName: file.name, sheetName, instructions, issues };
}

function buildModel(categoryImport, permissionImport) {
  const issues = [...categoryImport.issues, ...permissionImport.issues];
  const categories = categoryImport.categories;
  const categoryMap = new Map(categories.map(category => [category.categoryId, category]));
  const childrenMap = new Map();

  categories.forEach(category => {
    if (!childrenMap.has(category.parentId)) childrenMap.set(category.parentId, []);
    childrenMap.get(category.parentId).push(category);
  });
  childrenMap.forEach(children => children.sort((a, b) => categoryLabel(a).localeCompare(categoryLabel(b), 'ru')));
  const roots = childrenMap.get(null) || [];

  const leafCache = new Map();
  function leavesOf(categoryId) {
    if (leafCache.has(categoryId)) return leafCache.get(categoryId);
    const children = childrenMap.get(categoryId) || [];
    const leaves = children.length ? children.flatMap(child => leavesOf(child.categoryId)) : [categoryId];
    leafCache.set(categoryId, leaves);
    return leaves;
  }
  categories.forEach(category => leavesOf(category.categoryId));

  const byPartner = new Map();
  permissionImport.instructions.forEach(instruction => {
    if (!byPartner.has(instruction.partnerId)) byPartner.set(instruction.partnerId, []);
    byPartner.get(instruction.partnerId).push(instruction);
  });

  const partnerResults = new Map();
  byPartner.forEach((sourceRows, partnerId) => {
    const knownRows = [];
    sourceRows.forEach(row => {
      if (!categoryMap.has(row.categoryId)) {
        issues.push(makeIssue('error', 'Excel', `Категория ${row.categoryId} отсутствует в JSON. Сервер отклонит весь файл.`, {
          row: row.sourceRow,
          partnerId,
          categoryId: row.categoryId
        }));
      } else {
        knownRows.push(row);
      }
    });

    const batches = [];
    knownRows.forEach(row => {
      let batch = batches[batches.length - 1];
      if (!batch || batch.allow !== row.allow) {
        batch = { number: batches.length + 1, allow: row.allow, rows: [] };
        batches.push(batch);
      }
      batch.rows.push(row);
    });

    const allowed = new Set();
    batches.forEach(batch => {
      const rowLeaves = batch.rows.map(row => new Set(leavesOf(row.categoryId)));
      batch.rows.forEach((row, rowIndex) => {
        const otherCoverage = new Set();
        rowLeaves.forEach((leaves, otherIndex) => {
          if (otherIndex !== rowIndex) leaves.forEach(leaf => otherCoverage.add(leaf));
        });
        if (rowLeaves[rowIndex].size > 0 && [...rowLeaves[rowIndex]].every(leaf => otherCoverage.has(leaf))) {
          issues.push(makeIssue('warning', 'Excel', 'Избыточная инструкция: поддерево уже покрыто другим элементом этого батча.', {
            row: row.sourceRow,
            partnerId,
            categoryId: row.categoryId
          }));
        }
      });

      const coverage = new Set();
      rowLeaves.forEach(leaves => leaves.forEach(leaf => coverage.add(leaf)));
      batch.leafCount = coverage.size;
      coverage.forEach(leaf => batch.allow === 'yes' ? allowed.add(leaf) : allowed.delete(leaf));
    });

    partnerResults.set(partnerId, { partnerId, sourceRows, allowed, batches });
  });

  const partners = [...partnerResults.keys()].sort((a, b) => a.localeCompare(b));
  const statusesByPartner = new Map();

  partners.forEach(partnerId => {
    const statuses = new Map();
    const allowed = partnerResults.get(partnerId).allowed;
    function calculate(categoryId) {
      if (statuses.has(categoryId)) return statuses.get(categoryId);
      const children = childrenMap.get(categoryId) || [];
      let status;
      if (!children.length) {
        status = allowed.has(categoryId) ? 'allowed' : 'blocked';
      } else {
        const childStatuses = children.map(child => calculate(child.categoryId));
        status = childStatuses.every(value => value === 'allowed')
          ? 'allowed'
          : childStatuses.every(value => value === 'blocked')
            ? 'blocked'
            : 'partial';
      }
      statuses.set(categoryId, status);
      return status;
    }
    roots.forEach(root => calculate(root.categoryId));
    statusesByPartner.set(partnerId, statuses);
  });

  const aggregates = new Map();
  categories.forEach(category => {
    const aggregate = { allowed: 0, partial: 0, blocked: 0, restricted: 0 };
    partners.forEach(partnerId => aggregate[statusesByPartner.get(partnerId).get(category.categoryId)]++);
    aggregate.restricted = aggregate.partial + aggregate.blocked;
    aggregates.set(category.categoryId, aggregate);
  });

  const operationCounts = new Map();
  permissionImport.instructions.forEach(row => {
    const key = `${row.partnerId}:${row.categoryId}`;
    if (!operationCounts.has(key)) operationCounts.set(key, { yes: 0, no: 0 });
    operationCounts.get(key)[row.allow]++;
    const allKey = `*:${row.categoryId}`;
    if (!operationCounts.has(allKey)) operationCounts.set(allKey, { yes: 0, no: 0 });
    operationCounts.get(allKey)[row.allow]++;
  });

  return {
    categories,
    categoryMap,
    childrenMap,
    roots,
    leavesOf,
    partners,
    partnerResults,
    statusesByPartner,
    aggregates,
    operationCounts,
    issues
  };
}

function categoryLabel(category) {
  return category.displayName || category.name || `Категория ${category.categoryId}`;
}

async function loadCategoryFile(file) {
  state.loadingMessage = `Чтение ${file.name}...`;
  render();
  state.categoryImport = await parseCategoryFile(file);
  rebuildModel();
}

async function loadPermissionFile(file) {
  state.loadingMessage = `Чтение ${file.name}...`;
  render();
  state.permissionImport = await parsePermissionFile(file);
  rebuildModel();
}

function rebuildModel() {
  state.loadingMessage = '';
  state.model = null;
  state.issues = [
    ...(state.categoryImport?.issues || []),
    ...(state.permissionImport?.issues || [])
  ];

  if (state.categoryImport && state.permissionImport && state.categoryImport.structurallyValid) {
    state.model = buildModel(state.categoryImport, state.permissionImport);
    state.issues = state.model.issues;
    state.selectedCategoryId = state.model.roots[0]?.categoryId ?? null;
    state.selectedPartnerId = state.model.partners[0] ?? null;
    state.expanded = new Set(state.model.roots.map(category => category.categoryId));
    state.mode = 'categories';
    state.filter = 'all';
    state.search = '';
  }
  render();
}

function renderHeader() {
  const modelReady = Boolean(state.model);
  const errorCount = state.issues.filter(issue => issue.severity === 'error').length;
  const warningCount = state.issues.filter(issue => issue.severity === 'warning').length;
  const healthClass = errorCount ? 'has-errors' : warningCount ? 'has-warnings' : '';

  return `<header class="app-header">
    <div class="brand"><span class="brand-mark">±</span><span>Доступность категорий</span></div>
    ${renderHeaderFileInput('permission-file', '.xlsx,.xls', state.permissionImport?.fileName, 'Выбрать Excel')}
    ${renderHeaderFileInput('category-file', '.json,application/json', state.categoryImport?.fileName, 'Выбрать JSON')}
    <div class="header-spacer"></div>
    <button class="header-button" data-action="export" ${modelReady ? '' : 'disabled'} title="Скачать отчет Excel"><span aria-hidden="true">⇩</span><span>Экспорт XLSX</span></button>
  </header>
  <div class="workspace-bar">
    <nav class="tabs" aria-label="Режим просмотра">
      ${renderTab('categories', 'Категории', modelReady)}
      ${renderTab('partners', 'Партнеры', modelReady)}
      ${renderTab('quality', 'Качество данных', state.issues.length > 0 || modelReady)}
    </nav>
    <label class="search-wrap"><span class="search-icon" aria-hidden="true">⌕</span><input id="search" value="${escapeHtml(state.search)}" placeholder="Категория или ID" ${modelReady && state.mode !== 'quality' ? '' : 'disabled'}></label>
    <button class="health-button ${healthClass}" data-mode="quality" ${state.issues.length || modelReady ? '' : 'disabled'}><span class="dot"></span>${errorCount} ошибок · ${warningCount} предупреждений</button>
  </div>`;
}

function renderHeaderFileInput(id, accept, fileName, emptyLabel) {
  return `<label class="file-button ${fileName ? 'is-loaded' : ''}" title="${fileName ? 'Выбрать другой файл' : emptyLabel}">
    <input id="${id}" type="file" accept="${accept}">
    <span aria-hidden="true">${fileName ? '✓' : '+'}</span>
    <span class="file-name">${escapeHtml(fileName || emptyLabel)}</span>
  </label>`;
}

function renderTab(mode, label, enabled) {
  return `<button class="tab ${state.mode === mode ? 'is-active' : ''}" data-mode="${mode}" ${enabled ? '' : 'disabled'}>${label}</button>`;
}

function renderImportWorkspace() {
  const errors = state.issues.filter(issue => issue.severity === 'error');
  const statusClass = errors.length ? 'is-error' : '';
  let status = state.loadingMessage || 'Выберите оба файла. Обработка выполняется локально в браузере.';
  if (!state.loadingMessage && errors.length) status = `${errors.length} ошибок: ${errors.slice(0, 3).map(issue => issue.message).join(' · ')}`;

  return `<section class="import-workspace">
    <div class="import-panel">
      <div class="import-heading"><h1>Загрузка данных</h1><p>Первый лист Excel и дерево категорий будут проверены перед построением отчета.</p></div>
      <div class="import-files">
        ${renderImportFile('permission-file-empty', '.xlsx,.xls', state.permissionImport?.fileName, 'Excel с разрешениями', 'partnerId · categoryIds · allow')}
        ${renderImportFile('category-file-empty', '.json,application/json', state.categoryImport?.fileName, 'JSON с категориями', 'categoryId · parentid · displayname')}
      </div>
      <div class="import-status ${statusClass}">${escapeHtml(status)}</div>
    </div>
  </section>`;
}

function renderImportFile(id, accept, fileName, label, hint) {
  return `<label class="import-file ${fileName ? 'is-loaded' : ''}">
    <input id="${id}" type="file" accept="${accept}">
    <span class="import-icon" aria-hidden="true">${fileName ? '✓' : '+'}</span>
    <strong>${escapeHtml(fileName || label)}</strong>
    <span>${escapeHtml(fileName ? 'Нажмите, чтобы заменить файл' : hint)}</span>
  </label>`;
}

function statusLabel(status) {
  return { allowed: 'Разрешено', partial: 'Частично', blocked: 'Запрещено' }[status];
}

function getStatus(partnerId, categoryId) {
  return state.model.statusesByPartner.get(partnerId).get(categoryId);
}

function getOperationCounts(categoryId, partnerId = null) {
  return state.model.operationCounts.get(`${partnerId || '*'}:${categoryId}`) || { yes: 0, no: 0 };
}

function categoryMatcher() {
  const cache = new Map();
  const query = state.search.trim().toLowerCase();

  function directMatch(category) {
    const searchMatch = !query || `${categoryLabel(category)} ${category.categoryId}`.toLowerCase().includes(query);
    if (!searchMatch) return false;
    if (state.mode !== 'categories' || state.filter === 'all') return true;
    if (state.filter === 'restricted') return state.model.aggregates.get(category.categoryId).restricted > 0;
    return state.issues.some(issue => issue.severity === 'warning' && issue.categoryId === category.categoryId);
  }

  function subtreeMatch(category) {
    if (cache.has(category.categoryId)) return cache.get(category.categoryId);
    const matches = directMatch(category) || (state.model.childrenMap.get(category.categoryId) || []).some(subtreeMatch);
    cache.set(category.categoryId, matches);
    return matches;
  }

  return subtreeMatch;
}

function renderTreeRows(nodes, depth, partnerId, matches) {
  return nodes.map(category => {
    if (!matches(category)) return '';
    const categoryId = category.categoryId;
    const children = state.model.childrenMap.get(categoryId) || [];
    const expanded = state.search || state.filter !== 'all' || state.expanded.has(categoryId);
    const operations = getOperationCounts(categoryId, partnerId);
    const aggregate = state.model.aggregates.get(categoryId);
    const partnerStatus = partnerId ? getStatus(partnerId, categoryId) : null;

    const statusCell = partnerId
      ? `<span class="status status-${partnerStatus}"><span class="dot"></span>${statusLabel(partnerStatus)}</span>`
      : `<div class="restriction-cell"><span class="restriction-count">${aggregate.restricted}</span>${renderMiniBar(aggregate)}</div>`;

    let allowedLeaves = '';
    if (partnerId) {
      const result = state.model.partnerResults.get(partnerId);
      const leaves = state.model.leavesOf(categoryId);
      allowedLeaves = `${leaves.filter(leaf => result.allowed.has(leaf)).length}/${leaves.length}`;
    }

    return `<div class="tree-row ${state.selectedCategoryId === categoryId ? 'is-selected' : ''}" data-select-category="${categoryId}">
      <div class="tree-name">
        <span class="indent" style="width:${depth * 18}px"></span>
        ${children.length ? `<button class="chevron" data-toggle-category="${categoryId}" title="${expanded ? 'Свернуть' : 'Раскрыть'}">${expanded ? '⌄' : '›'}</button>` : '<span class="chevron-spacer"></span>'}
        <span class="category-label">${escapeHtml(categoryLabel(category))}</span><span class="category-id">#${categoryId}</span>
      </div>
      ${statusCell}
      <span class="muted-number">${partnerId ? allowedLeaves : aggregate.partial}</span>
      <span class="op-markers">${operations.yes ? `<span class="op-marker op-add" title="Инструкций allow = yes">YES ${operations.yes}</span>` : ''}${operations.no ? `<span class="op-marker op-remove" title="Инструкций allow = no">NO ${operations.no}</span>` : ''}</span>
    </div>${children.length && expanded ? renderTreeRows(children, depth + 1, partnerId, matches) : ''}`;
  }).join('');
}

function renderMiniBar(aggregate) {
  const total = state.model.partners.length || 1;
  return `<span class="mini-bar" title="Полностью: ${aggregate.blocked}; частично: ${aggregate.partial}"><span class="bar-blocked" style="width:${aggregate.blocked / total * 100}%"></span><span class="bar-partial" style="width:${aggregate.partial / total * 100}%"></span></span>`;
}

function renderAlert() {
  const errors = state.issues.filter(issue => issue.severity === 'error');
  if (!errors.length) return '';
  return `<div class="alert-banner"><span aria-hidden="true">!</span><strong>Сервер отклонит этот файл.</strong><span>${escapeHtml(errors[0].message)}${errors.length > 1 ? ` Еще ошибок: ${errors.length - 1}.` : ''}</span><button data-mode="quality">Открыть ошибки</button></div>`;
}

function renderSummary() {
  const rootRestrictions = state.model.roots.reduce((sum, category) => sum + state.model.aggregates.get(category.categoryId).restricted, 0);
  return `<div class="summary-strip">
    <div class="summary-item"><span class="summary-value">${state.model.categories.length}</span><span class="summary-label">категорий</span></div>
    <div class="summary-item"><span class="summary-value">${state.model.partners.length}</span><span class="summary-label">партнеров</span></div>
    <div class="summary-item"><span class="summary-value">${state.permissionImport.instructions.length}</span><span class="summary-label">валидных инструкций</span></div>
    <div class="summary-item"><span class="summary-value">${rootRestrictions}</span><span class="summary-label">ограничений в верхних ветках</span></div>
  </div>`;
}

function renderFilter() {
  return `<div class="segmented" aria-label="Фильтр дерева">
    ${renderSegment('all', 'Все')}${renderSegment('restricted', 'С ограничениями')}${renderSegment('warnings', 'Предупреждения')}
  </div>`;
}

function renderSegment(filter, label) {
  return `<button class="segment ${state.filter === filter ? 'is-active' : ''}" data-filter="${filter}">${label}</button>`;
}

function renderPartnerSelect() {
  return `<select id="partner-select" class="partner-select" aria-label="Выбрать партнера">${state.model.partners.map(partnerId => `<option value="${escapeHtml(partnerId)}" ${partnerId === state.selectedPartnerId ? 'selected' : ''}>${escapeHtml(partnerId)}</option>`).join('')}</select>`;
}

function renderInspector() {
  const category = state.model.categoryMap.get(state.selectedCategoryId);
  if (!category) return '<aside class="inspector"></aside>';
  const categoryId = category.categoryId;
  const aggregate = state.model.aggregates.get(categoryId);
  const sortedPartners = [...state.model.partners].sort((first, second) => {
    const order = { blocked: 0, partial: 1, allowed: 2 };
    return order[getStatus(first, categoryId)] - order[getStatus(second, categoryId)] || first.localeCompare(second);
  });
  const selectedResult = state.model.partnerResults.get(state.selectedPartnerId);

  return `<aside class="inspector">
    <div class="inspector-header"><div class="eyebrow">Категория #${categoryId}</div><h2>${escapeHtml(categoryLabel(category))}</h2><div class="inspector-subtitle">${state.model.leavesOf(categoryId).length} товарных категорий в ветке</div></div>
    <div class="inspector-metrics">
      <div class="metric"><b>${aggregate.restricted}</b><span>с ограничениями</span></div>
      <div class="metric"><b>${aggregate.blocked}</b><span>полностью</span></div>
      <div class="metric"><b>${aggregate.partial}</b><span>частично</span></div>
    </div>
    <div class="inspector-body">
      <div class="subhead">Партнеры <span>${state.model.partners.length}</span></div>
      <div class="partner-list">${sortedPartners.map(partnerId => {
        const status = getStatus(partnerId, categoryId);
        return `<button class="partner-item" data-select-partner="${escapeHtml(partnerId)}"><span class="partner-code">${escapeHtml(partnerId)}</span><span class="status status-${status}"><span class="dot"></span>${statusLabel(status)}</span></button>`;
      }).join('')}</div>
      ${selectedResult ? `<details class="history"><summary>История батчей · ${escapeHtml(state.selectedPartnerId)}</summary>${selectedResult.batches.map(batch => `<div class="batch-row"><span>Батч ${batch.number}</span><span class="batch-op ${batch.allow === 'yes' ? 'op-add' : 'op-remove'}">${batch.allow.toUpperCase()}</span><span>${batch.rows.map(row => `${escapeHtml(categoryLabel(state.model.categoryMap.get(row.categoryId)))} (#${row.categoryId})`).join(', ')}</span></div>`).join('')}</details>` : ''}
    </div>
  </aside>`;
}

function renderTreeWorkspace() {
  const partnerMode = state.mode === 'partners';
  const partnerId = partnerMode ? state.selectedPartnerId : null;
  const matches = categoryMatcher();
  const title = partnerMode ? escapeHtml(partnerId) : 'Все категории';
  const alert = renderAlert();

  return `<div class="main-layout ${alert ? 'has-alert' : ''}">
    ${alert}
    <section class="tree-pane">
      <div class="section-toolbar">
        <div class="section-title"><h2>${title}</h2><p>${partnerMode ? 'Итоговый whitelist партнера' : 'Ограничения по всем партнерам'}</p></div>
        <div class="toolbar-spacer"></div>
        ${partnerMode ? renderPartnerSelect() : renderFilter()}
      </div>
      ${renderSummary()}
      <div class="tree-scroll">
        <div class="tree-head"><span>Категория</span><span>${partnerMode ? 'Состояние' : 'С ограничениями'}</span><span>${partnerMode ? 'Листья' : 'Частично'}</span><span>Инструкции</span></div>
        ${renderTreeRows(state.model.roots, 0, partnerId, matches)}
      </div>
    </section>
    ${renderInspector()}
  </div>`;
}

function renderIssueRows(issues) {
  if (!issues.length) return '<tr><td class="empty-row" colspan="6">Нет записей</td></tr>';
  return issues.map(issue => `<tr><td>${escapeHtml(issue.source)}</td><td>${issue.row ?? ''}</td><td>${escapeHtml(issue.partnerId || '')}</td><td>${issue.categoryId ?? ''}</td><td>${escapeHtml(issue.message)}</td></tr>`).join('');
}

function renderQuality() {
  const errors = state.issues.filter(issue => issue.severity === 'error');
  const warnings = state.issues.filter(issue => issue.severity === 'warning');
  return `<div class="quality-layout">
    <aside class="quality-summary"><h2>Качество данных</h2><div class="quality-number"><b style="color:var(--red)">${errors.length}</b><span>блокирующих ошибок</span></div><div class="quality-number"><b style="color:var(--amber)">${warnings.length}</b><span>предупреждений</span></div><div class="quality-number"><b>${state.permissionImport?.instructions.length || 0}</b><span>валидных инструкций</span></div></aside>
    <div class="quality-scroll"><div class="issue-group">
      <h3>Блокирующие ошибки</h3><table class="issue-table"><thead><tr><th>Источник</th><th>Строка</th><th>Партнер</th><th>Категория</th><th>Проблема</th></tr></thead><tbody>${renderIssueRows(errors)}</tbody></table>
      <h3>Предупреждения</h3><table class="issue-table"><thead><tr><th>Источник</th><th>Строка</th><th>Партнер</th><th>Категория</th><th>Проблема</th></tr></thead><tbody>${renderIssueRows(warnings)}</tbody></table>
    </div></div>
  </div>`;
}

function rowsToSheet(rows, headers) {
  if (!rows.length) return XLSX.utils.aoa_to_sheet([headers]);
  return XLSX.utils.json_to_sheet(rows, { header: headers });
}

function exportReport() {
  if (!state.model) return;
  const categoryRows = state.model.categories.map(category => {
    const aggregate = state.model.aggregates.get(category.categoryId);
    return {
      categoryId: category.categoryId,
      parentId: category.parentId,
      displayName: categoryLabel(category),
      leafCount: state.model.leavesOf(category.categoryId).length,
      partnersWithRestrictions: aggregate.restricted,
      fullyBlockedPartners: aggregate.blocked,
      partiallyRestrictedPartners: aggregate.partial
    };
  });

  const allLeafIds = new Set(state.model.categories.filter(category => !(state.model.childrenMap.get(category.categoryId) || []).length).map(category => category.categoryId));
  const partnerRows = state.model.partners.map(partnerId => {
    const result = state.model.partnerResults.get(partnerId);
    return {
      partnerId,
      allowedLeafCategories: [...result.allowed].filter(id => allLeafIds.has(id)).length,
      totalLeafCategories: allLeafIds.size,
      batches: result.batches.length,
      sourceInstructions: result.sourceRows.length
    };
  });

  const issueRows = issues => issues.map(issue => ({
    source: issue.source,
    row: issue.row ?? '',
    partnerId: issue.partnerId ?? '',
    categoryId: issue.categoryId ?? '',
    message: issue.message
  }));
  const errors = issueRows(state.issues.filter(issue => issue.severity === 'error'));
  const warnings = issueRows(state.issues.filter(issue => issue.severity === 'warning'));

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, rowsToSheet(categoryRows, ['categoryId', 'parentId', 'displayName', 'leafCount', 'partnersWithRestrictions', 'fullyBlockedPartners', 'partiallyRestrictedPartners']), 'Категории');
  XLSX.utils.book_append_sheet(workbook, rowsToSheet(partnerRows, ['partnerId', 'allowedLeafCategories', 'totalLeafCategories', 'batches', 'sourceInstructions']), 'Партнеры');
  XLSX.utils.book_append_sheet(workbook, rowsToSheet(errors, ['source', 'row', 'partnerId', 'categoryId', 'message']), 'Ошибки');
  XLSX.utils.book_append_sheet(workbook, rowsToSheet(warnings, ['source', 'row', 'partnerId', 'categoryId', 'message']), 'Предупреждения');
  XLSX.writeFile(workbook, 'category-permissions-report.xlsx');
}

function render() {
  let workspace;
  if (!state.model) workspace = state.mode === 'quality' && state.issues.length ? renderQuality() : renderImportWorkspace();
  else workspace = state.mode === 'quality' ? renderQuality() : renderTreeWorkspace();
  rootElement.innerHTML = `<div class="app-shell">${renderHeader()}<main class="workspace">${workspace}</main></div>`;
}

document.addEventListener('change', async event => {
  const target = event.target;
  if (target.id === 'category-file' || target.id === 'category-file-empty') {
    if (target.files[0]) await loadCategoryFile(target.files[0]);
  } else if (target.id === 'permission-file' || target.id === 'permission-file-empty') {
    if (target.files[0]) await loadPermissionFile(target.files[0]);
  } else if (target.id === 'partner-select') {
    state.selectedPartnerId = target.value;
    render();
  }
});

document.addEventListener('click', event => {
  const target = event.target.closest('button, [data-select-category]');
  if (!target || target.disabled) return;

  if (target.dataset.mode) {
    state.mode = target.dataset.mode;
    if (state.mode !== 'categories') state.filter = 'all';
    render();
  } else if (target.dataset.filter) {
    state.filter = target.dataset.filter;
    render();
  } else if (target.dataset.toggleCategory) {
    const categoryId = Number(target.dataset.toggleCategory);
    state.expanded.has(categoryId) ? state.expanded.delete(categoryId) : state.expanded.add(categoryId);
    render();
  } else if (target.dataset.selectCategory) {
    state.selectedCategoryId = Number(target.dataset.selectCategory);
    render();
  } else if (target.dataset.selectPartner) {
    state.selectedPartnerId = target.dataset.selectPartner;
    state.mode = 'partners';
    state.filter = 'all';
    render();
  } else if (target.dataset.action === 'export') {
    exportReport();
  }
});

document.addEventListener('input', event => {
  if (event.target.id !== 'search') return;
  state.search = event.target.value;
  render();
  const input = document.getElementById('search');
  input.focus();
  input.setSelectionRange(state.search.length, state.search.length);
});

render();
