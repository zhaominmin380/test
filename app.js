(() => {
  const workbook = window.GRIDBOOK;
  const FORM_SHEETS = new Set(['銀行轉帳', '超商轉帳', '虛擬帳號', '其他入帳']);
  const PAYMENT_ROWS = [2, 3, 4, 5, 6];
  const SUMMARY_ROWS = [13, 14, 15, 16, 17];
  const amountColumns = Array.from({ length: 10 }, (_, index) => index + 4);
  const HISTORY_STANDARD_EDITABLE_COLUMNS = new Set([12, 14, 16, 17, 25, 27]);
  const HISTORY_FILTER_COLUMNS = [
    { column: 4, label: '委託人編號', searchable: true },
    { column: 5, label: '遊戲平台', searchable: false },
    { column: 6, label: '經營角色', searchable: false },
    { column: 7, label: '遊戲暱稱', searchable: true },
  ];
  const grid = document.querySelector('#sheet-grid');
  const sheetScroll = document.querySelector('#sheet-scroll');
  const tabs = document.querySelector('#sheet-tabs');
  const nameBox = document.querySelector('#name-box');
  const formulaContent = document.querySelector('#formula-content');
  const statusBar = document.querySelector('#status-bar');
  const hiddenButton = document.querySelector('#hidden-sheet-button');
  const hiddenMenu = document.querySelector('#hidden-sheet-menu');
  const horizontalScrollControl = document.querySelector('.horizontal-scroll-control');
  const horizontalScrollLeft = document.querySelector('#horizontal-scroll-left');
  const horizontalScrollRight = document.querySelector('#horizontal-scroll-right');
  const horizontalScrollTrack = document.querySelector('#horizontal-scroll-track');
  const horizontalScrollThumb = document.querySelector('#horizontal-scroll-thumb');
  const historyFilterMenu = document.querySelector('#history-filter-menu');
  const workbookLockButton = document.querySelector('#workbook-lock-button');
  const changePasswordButton = document.querySelector('#change-password-button');
  const lockDialog = document.querySelector('#lock-dialog');
  const lockDialogForm = document.querySelector('#lock-dialog-form');
  const lockDialogTitle = document.querySelector('#lock-dialog-title');
  const lockDialogDescription = document.querySelector('#lock-dialog-description');
  const lockDialogSubmit = document.querySelector('#lock-dialog-submit');
  const lockDialogCancel = document.querySelector('#lock-dialog-cancel');
  const lockDialogError = document.querySelector('#lock-dialog-error');
  const currentPasswordField = document.querySelector('#current-password-field');
  const currentPassword = document.querySelector('#current-password');
  const newPasswordField = document.querySelector('#new-password-field');
  const newPassword = document.querySelector('#new-password');
  const confirmPasswordField = document.querySelector('#confirm-password-field');
  const confirmPassword = document.querySelector('#confirm-password');
  const storageKey = 'v20-12-editable-workbook-v1';
  const lockStorageKey = 'v20-12-workbook-lock-v1';

  workbook.sheets = workbook.sheets.filter((sheet) => sheet.name !== '發票');
  try {
    const saved = JSON.parse(localStorage.getItem(storageKey) || 'null');
    if (saved?.sheets?.length) workbook.sheets = saved.sheets.filter((sheet) => sheet.name !== '發票');
  } catch {
    localStorage.removeItem(storageKey);
  }

  let activeSheet = workbook.sheets.find((sheet) => sheet.state === 'visible') || workbook.sheets[0];
  let activeCell = null;
  let selectionAnchor = null;
  let selectionFocus = null;
  let sheetCanvasWidth = 0;
  let sequence = 0;
  let statusNotice = '';
  let lockDialogMode = 'unlock';
  let workbookUnlocked = false;
  const historyFilters = new Map();
  let activeHistoryFilterMenu = null;
  let lockConfig = { passwordHash: '' };
  const renderedCellElements = new Map();
  const selectedCellElements = new Set();

  const cellKey = (row, column) => `${row}:${column}`;
  const columnNumber = (label) => {
    let value = 0;
    for (const character of label) value = value * 26 + character.charCodeAt(0) - 64;
    return value;
  };
  const rowNumber = (label) => Number(label.match(/\d+/)?.[0] || 1);
  const columnLabel = (label) => label.match(/[A-Z]+/)?.[0] || 'A';
  const getSheet = (name) => workbook.sheets.find((sheet) => sheet.name === name);
  const columnLabelFor = (column) => {
    let value = column;
    let label = '';
    while (value > 0) {
      value -= 1;
      label = String.fromCharCode(65 + value % 26) + label;
      value = Math.floor(value / 26);
    }
    return label;
  };

  function syncHorizontalScrollControl() {
    const maximum = Math.max(0, Math.ceil(sheetScroll.scrollWidth - sheetScroll.clientWidth));
    const trackWidth = horizontalScrollTrack.clientWidth;
    const thumbWidth = maximum === 0 ? trackWidth : Math.max(28, Math.min(trackWidth, Math.round(trackWidth * sheetScroll.clientWidth / sheetScroll.scrollWidth)));
    const thumbTravel = Math.max(0, trackWidth - thumbWidth);
    const thumbOffset = maximum === 0 ? 0 : Math.round(thumbTravel * sheetScroll.scrollLeft / maximum);
    horizontalScrollThumb.style.width = `${thumbWidth}px`;
    horizontalScrollThumb.style.transform = `translateX(${thumbOffset}px)`;
    horizontalScrollTrack.setAttribute('aria-valuemax', String(maximum));
    horizontalScrollTrack.setAttribute('aria-valuenow', String(Math.min(maximum, Math.round(sheetScroll.scrollLeft))));
    horizontalScrollLeft.disabled = maximum === 0 || sheetScroll.scrollLeft <= 0;
    horizontalScrollRight.disabled = maximum === 0 || sheetScroll.scrollLeft >= maximum;
    horizontalScrollControl.classList.toggle('is-disabled', maximum === 0);
  }

  function moveHorizontalScroll(value) {
    const maximum = Math.max(0, sheetScroll.scrollWidth - sheetScroll.clientWidth);
    sheetScroll.scrollLeft = Math.min(maximum, Math.max(0, value));
    syncHorizontalScrollControl();
  }

  function scrollFromTrackPosition(clientX, centerThumb = true) {
    const maximum = Math.max(0, sheetScroll.scrollWidth - sheetScroll.clientWidth);
    const bounds = horizontalScrollTrack.getBoundingClientRect();
    const thumbWidth = horizontalScrollThumb.getBoundingClientRect().width;
    const travel = Math.max(0, bounds.width - thumbWidth);
    const rawPosition = clientX - bounds.left - (centerThumb ? thumbWidth / 2 : 0);
    const position = Math.min(travel, Math.max(0, rawPosition));
    moveHorizontalScroll(travel ? maximum * position / travel : 0);
  }

  function revealInvoiceAmountColumn() {
    const history = getSheet('歷史資料');
    const invoiceAmountIndex = history?.cells.find((cell) => cell.r === 2 && cell.v === '發票金額')?.c;
    const invoiceAmountColumn = invoiceAmountIndex ? history.columns[invoiceAmountIndex - 1] : null;
    if (invoiceAmountColumn) invoiceAmountColumn.hidden = false;
  }

  function getCell(sheet, row, column) {
    return sheet.cells.find((cell) => cell.r === row && cell.c === column);
  }

  function ensureHistoryDifferenceColumn() {
    const history = getSheet('歷史資料');
    if (!history) return;
    const existing = history.cells.find((cell) => cell.r === 2 && cell.v === '差異');
    if (!existing) {
      const differenceColumn = 13;
      history.columns.splice(differenceColumn - 1, 0, { letter: 'M', width: 46, hidden: false, outline: 0 });
      history.cells.forEach((cell) => {
        if (cell.c >= differenceColumn) cell.c += 1;
      });
      history.merges.forEach((merge) => {
        if (merge[1] >= differenceColumn) merge[1] += 1;
        if (merge[3] >= differenceColumn) merge[3] += 1;
      });
      history.maxCol = history.columns.length;
      const headerStyle = getCell(history, 2, 12)?.s || null;
      setCell(history, 2, differenceColumn, '差異', { style: headerStyle });
    }
    history.columns.forEach((column, index) => { column.letter = columnLabelFor(index + 1); });
    history.maxCol = history.columns.length;
  }

  function removeHistorySummaryColumns() {
    const history = getSheet('歷史資料');
    if (!history) return;
    const summaryColumn = history.cells.find((cell) => cell.r === 1 && cell.v === '收款總計')?.c;
    if (!summaryColumn) return;
    history.cells = history.cells.filter((cell) => cell.c < summaryColumn);
    history.columns.splice(summaryColumn - 1);
    history.merges = history.merges
      .filter((merge) => merge[1] < summaryColumn)
      .map((merge) => [merge[0], merge[1], merge[2], Math.min(merge[3], summaryColumn - 1)]);
    history.columns.forEach((column, index) => { column.letter = columnLabelFor(index + 1); });
    history.maxCol = history.columns.length;
  }

  ensureHistoryDifferenceColumn();
  removeHistorySummaryColumns();

  function getRaw(sheet, row, column) {
    const cell = getCell(sheet, row, column);
    return cell?.raw ?? cell?.v ?? '';
  }

  function numberValue(value) {
    if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
    const source = String(value ?? '').trim();
    if (!source) return 0;
    const percent = source.endsWith('%');
    const parsed = Number(source.replace(/,/g, '').replace('%', ''));
    if (!Number.isFinite(parsed)) return 0;
    return percent ? parsed / 100 : parsed;
  }

  function formatValue(raw, style) {
    if (raw === null || raw === undefined || raw === '') return '';
    if (typeof raw === 'boolean') return raw ? 'TRUE' : 'FALSE';
    if (typeof raw === 'string') return raw;
    const format = style?.numberFormat || 'General';
    if (format.includes('%')) {
      const match = format.match(/0\.(0+)%/);
      return `${(raw * 100).toFixed(match ? match[1].length : 0)}%`;
    }
    const match = format.match(/0\.(0+)/);
    const decimals = match ? match[1].length : 0;
    const grouping = format.includes(',');
    if (decimals) return raw.toLocaleString('en-US', { minimumFractionDigits: decimals, maximumFractionDigits: decimals, useGrouping: grouping });
    if (Number.isInteger(raw)) return raw.toLocaleString('en-US', { useGrouping: grouping });
    return String(raw);
  }

  function makeFillReadableWithBlackText(fill) {
    const channels = fill.match(/[\da-f]{2}/gi);
    if (!channels || channels.length !== 3) return fill;
    const rgb = channels.map((channel) => Number.parseInt(channel, 16));
    const brightness = (rgb[0] * 299 + rgb[1] * 587 + rgb[2] * 114) / 1000;
    if (brightness >= 145) return fill;
    return `#${rgb.map((channel) => Math.round(channel + (255 - channel) * 0.78).toString(16).padStart(2, '0')).join('')}`;
  }

  function sectionFill(sheet, row, column) {
    if (FORM_SHEETS.has(sheet.name)) {
      if (row === 1 && column >= 4 && column <= 8) return '#47D359';
      if (row >= 2 && column >= 4 && column <= 8) return '#C1F0C8';
      if (row === 1 && column >= 9 && column <= 13) return '#F1A983';
      if (row >= 2 && column >= 9 && column <= 13) return '#FBE2D5';
      if (row === 1 && column >= 14 && column <= 16) return '#47D359';
      if (row >= 2 && column >= 14 && column <= 16) return '#C1F0C8';
    }
    if (sheet.name === '歷史資料') {
      // Seq / 交易序號 are a separate neutral key column, not part of the blue basic-data block.
      if (column === 1 && (row === 1 || row === 2)) return '#D0D0D0';
      if (row === 1 && column >= 2 && column <= 10) return '#4D93D9';
      if (row === 2 && column >= 2 && column <= 10) return '#A6C9EC';
    }
    return null;
  }

  function setCell(sheet, row, column, raw, options = {}) {
    let cell = getCell(sheet, row, column);
    if (!cell) {
      cell = { r: row, c: column, raw: '', v: '', s: options.style || null };
      sheet.cells.push(cell);
    }
    if (options.style) cell.s = options.style;
    cell.raw = raw;
    cell.v = formatValue(raw, sheet.styles[cell.s]);
    if (options.checkbox !== undefined) cell.checkbox = options.checkbox;
    if (options.checked !== undefined) cell.checked = options.checked;
    return cell;
  }

  function removeCell(sheet, row, column) {
    sheet.cells = sheet.cells.filter((cell) => cell.r !== row || cell.c !== column);
  }

  function ensureRows(sheet, total) {
    while (sheet.rows.length < total) sheet.rows.push({ height: 20, hidden: false, outline: 0 });
    sheet.maxRow = Math.max(sheet.maxRow, total);
  }

  function ensureWorkbookCanvas() {
    workbook.sheets
      .filter((sheet) => sheet.name !== 'secret')
      .forEach((sheet) => ensureRows(sheet, 50));
  }

  function displayColumnsFor(sheet, rowHeaderWidth) {
    const columns = sheet.columns.map((column) => ({ ...column, virtual: false }));
    const usedWidth = columns.reduce((total, column) => total + (column.hidden ? 0 : column.width), 0);
    const availableWidth = Math.max(0, sheetScroll.clientWidth - rowHeaderWidth);
    const spacerWidth = 75;
    const spacerCount = Math.max(0, Math.ceil((availableWidth - usedWidth) / spacerWidth));
    for (let index = 0; index < spacerCount; index += 1) {
      columns.push({ letter: columnLabelFor(columns.length + 1), width: spacerWidth, hidden: false, virtual: true });
    }
    return columns;
  }

  function formatNow() {
    const now = new Date();
    const pad = (value) => String(value).padStart(2, '0');
    return `${now.getFullYear()}/${pad(now.getMonth() + 1)}/${pad(now.getDate())} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
  }

  function saveState() {
    localStorage.setItem(storageKey, JSON.stringify({ sheets: workbook.sheets }));
  }

  function historyRowHasData(sheet, row) {
    return Array.from({ length: 28 }, (_, index) => getRaw(sheet, row, index + 1))
      .some((value) => String(value ?? '').trim() !== '');
  }

  function historyCellValue(sheet, row, column) {
    const cell = getCell(sheet, row, column);
    return String(cell?.v ?? cell?.raw ?? '').trim();
  }

  function historyFilterConfig(column) {
    return HISTORY_FILTER_COLUMNS.find((filter) => filter.column === column);
  }

  function historyFilterValues(column) {
    const history = getSheet('歷史資料');
    if (!history) return [];
    return [...new Set(Array.from({ length: Math.max(0, history.maxRow - 2) }, (_, index) => index + 3)
      .filter((row) => historyRowHasData(history, row))
      .map((row) => historyCellValue(history, row, column))
      .filter(Boolean))]
      .sort((left, right) => left.localeCompare(right, 'zh-Hant', { numeric: true }));
  }

  function historyRowMatchesFilters(sheet, row) {
    if (!historyRowHasData(sheet, row)) return true;
    return [...historyFilters.entries()].every(([column, selectedValues]) => selectedValues.has(historyCellValue(sheet, row, column)));
  }

  function isHistoryRowVisible(sheet, row) {
    return sheet.name !== '歷史資料' || row < 3 || historyRowMatchesFilters(sheet, row);
  }

  function closeHistoryFilterMenu() {
    activeHistoryFilterMenu = null;
    historyFilterMenu.hidden = true;
    historyFilterMenu.replaceChildren();
  }

  function positionHistoryFilterMenu(anchor) {
    const bounds = anchor.getBoundingClientRect();
    const menuBounds = historyFilterMenu.getBoundingClientRect();
    const menuWidth = menuBounds.width || 270;
    const menuHeight = menuBounds.height || 280;
    const viewportWidth = window.innerWidth || 1024;
    const viewportHeight = window.innerHeight || 768;
    const left = Math.max(8, Math.min(bounds.right - menuWidth, viewportWidth - menuWidth - 8));
    const top = Math.max(8, Math.min(bounds.bottom + 4, viewportHeight - menuHeight - 8));
    historyFilterMenu.style.left = `${left}px`;
    historyFilterMenu.style.top = `${top}px`;
  }

  function renderHistoryFilterOptions() {
    if (!activeHistoryFilterMenu) return;
    const { values, query, draft, options } = activeHistoryFilterMenu;
    options.replaceChildren();
    const normalizedQuery = query.trim().toLocaleLowerCase();
    const visibleValues = values.filter((value) => value.toLocaleLowerCase().includes(normalizedQuery));
    if (!visibleValues.length) {
      const empty = document.createElement('span');
      empty.className = 'history-filter-empty';
      empty.textContent = '找不到符合的值';
      options.append(empty);
      return;
    }
    visibleValues.forEach((value) => {
      const option = document.createElement('label');
      option.className = 'history-filter-option';
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.checked = draft.has(value);
      checkbox.addEventListener('change', () => {
        if (checkbox.checked) draft.add(value);
        else draft.delete(value);
      });
      const text = document.createElement('span');
      text.textContent = value;
      option.append(checkbox, text);
      options.append(option);
    });
  }

  function openHistoryFilterMenu(column, anchor) {
    const filter = historyFilterConfig(column);
    if (!filter || activeSheet?.name !== '歷史資料') return;
    if (activeHistoryFilterMenu?.column === column) {
      closeHistoryFilterMenu();
      return;
    }
    const values = historyFilterValues(column);
    const selectedValues = historyFilters.get(column);
    const draft = new Set(selectedValues || values);
    const title = document.createElement('p');
    title.className = 'history-filter-menu-title';
    title.textContent = `篩選：${filter.label}`;
    const options = document.createElement('div');
    options.className = 'history-filter-options';
    activeHistoryFilterMenu = { column, anchor, values, draft, options, query: '' };
    historyFilterMenu.replaceChildren(title);
    if (filter.searchable) {
      const search = document.createElement('input');
      search.className = 'history-filter-search';
      search.type = 'search';
      search.placeholder = `搜尋${filter.label}`;
      search.autocomplete = 'off';
      search.addEventListener('input', () => {
        activeHistoryFilterMenu.query = search.value;
        renderHistoryFilterOptions();
      });
      historyFilterMenu.append(search);
    }
    historyFilterMenu.append(options);
    const actions = document.createElement('div');
    actions.className = 'history-filter-actions';
    const clear = document.createElement('button');
    clear.type = 'button';
    clear.textContent = '清除篩選';
    clear.addEventListener('click', () => {
      historyFilters.delete(column);
      closeHistoryFilterMenu();
      renderGrid(activeSheet);
    });
    const selectAll = document.createElement('button');
    selectAll.type = 'button';
    selectAll.textContent = '全選';
    selectAll.addEventListener('click', () => {
      activeHistoryFilterMenu.draft = new Set(values);
      renderHistoryFilterOptions();
    });
    const selectNone = document.createElement('button');
    selectNone.type = 'button';
    selectNone.textContent = '全不選';
    selectNone.addEventListener('click', () => {
      activeHistoryFilterMenu.draft.clear();
      renderHistoryFilterOptions();
    });
    const apply = document.createElement('button');
    apply.type = 'button';
    apply.className = 'filter-apply';
    apply.textContent = '套用';
    apply.addEventListener('click', () => {
      const allSelected = values.every((value) => activeHistoryFilterMenu.draft.has(value));
      if (allSelected) historyFilters.delete(column);
      else historyFilters.set(column, new Set(activeHistoryFilterMenu.draft));
      closeHistoryFilterMenu();
      renderGrid(activeSheet);
    });
    actions.append(clear, selectAll, selectNone, apply);
    historyFilterMenu.append(actions);
    historyFilterMenu.hidden = false;
    renderHistoryFilterOptions();
    positionHistoryFilterMenu(anchor);
    requestAnimationFrame(() => positionHistoryFilterMenu(anchor));
  }

  function renderHistoryFilterButton(sheet, row, column) {
    if (sheet.name !== '歷史資料' || row !== 2 || !historyFilterConfig(column)) return null;
    const button = document.createElement('button');
    button.className = `column-filter-button${historyFilters.has(column) ? ' is-active' : ''}`;
    button.type = 'button';
    button.textContent = '▼';
    button.title = `篩選${historyFilterConfig(column).label}`;
    button.setAttribute('aria-label', `篩選${historyFilterConfig(column).label}`);
    button.addEventListener('click', (event) => {
      event.preventDefault();
      event.stopPropagation();
      openHistoryFilterMenu(column, button);
    });
    return button;
  }

  async function hashPassword(value) {
    if (window.crypto?.subtle && window.TextEncoder) {
      const source = new TextEncoder().encode(value);
      const digest = await window.crypto.subtle.digest('SHA-256', source);
      return Array.from(new Uint8Array(digest), (part) => part.toString(16).padStart(2, '0')).join('');
    }
    let hash = 2166136261;
    for (const character of value) {
      hash ^= character.charCodeAt(0);
      hash = Math.imul(hash, 16777619);
    }
    return `fallback-${(hash >>> 0).toString(16)}`;
  }

  function updateLockControls() {
    workbookLockButton.textContent = workbookUnlocked ? '🔓 已解鎖（按此上鎖）' : '🔒 已上鎖';
    workbookLockButton.setAttribute('aria-pressed', String(!workbookUnlocked));
    changePasswordButton.hidden = !workbookUnlocked;
    changePasswordButton.disabled = !workbookUnlocked;
  }

  function canOpenSheet(sheet) {
    return sheet.name !== '設定' || workbookUnlocked;
  }

  async function initializeWorkbookLock() {
    try {
      const savedLock = JSON.parse(localStorage.getItem(lockStorageKey) || 'null');
      if (typeof savedLock?.passwordHash === 'string' && savedLock.passwordHash) {
        lockConfig = savedLock;
      }
    } catch {
      lockConfig = { passwordHash: '' };
    }
    updateLockControls();
  }

  function showLockDialogError(message = '') {
    lockDialogError.textContent = message;
    lockDialogError.hidden = !message;
  }

  async function openLockDialog(mode) {
    await lockReady;
    lockDialogMode = mode;
    currentPassword.value = '';
    newPassword.value = '';
    confirmPassword.value = '';
    showLockDialogError();
    const settingPassword = mode === 'setup-password' || mode === 'change-password';
    const initialSetup = mode === 'setup-password';
    lockDialogTitle.textContent = initialSetup ? '設定工作簿密碼' : (settingPassword ? '更改工作簿密碼' : '解除工作簿鎖定');
    lockDialogDescription.textContent = initialSetup
      ? '請設定至少 4 碼的密碼；設定後工作簿將立即解鎖。'
      : (settingPassword ? '請先驗證目前密碼，再設定至少 4 碼的新密碼。' : '輸入密碼後，即可編輯所有資料欄位。');
    currentPasswordField.hidden = initialSetup;
    currentPassword.required = !initialSetup;
    newPasswordField.hidden = !settingPassword;
    confirmPasswordField.hidden = !settingPassword;
    newPassword.required = settingPassword;
    confirmPassword.required = settingPassword;
    lockDialogSubmit.textContent = settingPassword ? '儲存密碼' : '解鎖';
    if (!lockDialog.open) lockDialog.showModal();
    (initialSetup ? newPassword : currentPassword).focus();
  }

  async function handleLockDialogSubmit(event) {
    event.preventDefault();
    await lockReady;
    const initialSetup = lockDialogMode === 'setup-password';
    if (!initialSetup && await hashPassword(currentPassword.value) !== lockConfig.passwordHash) {
      showLockDialogError('目前密碼不正確。');
      currentPassword.focus();
      return;
    }
    if (lockDialogMode === 'change-password' || initialSetup) {
      if (newPassword.value.length < 4) {
        showLockDialogError('新密碼至少需要 4 碼。');
        newPassword.focus();
        return;
      }
      if (newPassword.value !== confirmPassword.value) {
        showLockDialogError('兩次輸入的新密碼不一致。');
        confirmPassword.focus();
        return;
      }
      lockConfig = { passwordHash: await hashPassword(newPassword.value) };
      localStorage.setItem(lockStorageKey, JSON.stringify(lockConfig));
      if (initialSetup) {
        workbookUnlocked = true;
        statusNotice = '工作簿密碼已設定，現在可編輯所有資料欄位。';
      } else {
        statusNotice = '工作簿密碼已更新。';
      }
    } else {
      workbookUnlocked = true;
      statusNotice = '工作簿已解鎖，所有資料欄位皆可雙擊修改。';
    }
    lockDialog.close();
    updateLockControls();
    renderGrid(activeSheet);
    renderTabs();
  }

  function freezePane(sheet) {
    if (!sheet.freeze) return { rows: 0, columns: 0 };
    return { rows: Math.max(0, rowNumber(sheet.freeze) - 1), columns: Math.max(0, columnNumber(columnLabel(sheet.freeze)) - 1) };
  }

  function mergeIndex(sheet) {
    const index = new Map();
    sheet.merges.forEach(([top, left, bottom, right]) => {
      for (let row = top; row <= bottom; row += 1) {
        for (let column = left; column <= right; column += 1) {
          index.set(cellKey(row, column), { top, left, rowSpan: bottom - top + 1, columnSpan: right - left + 1 });
        }
      }
    });
    return index;
  }

  function applyStyle(element, style, fillOverride = null) {
    // Newly entered blank cells have no source style, but still belong to a coloured section.
    if (!style) {
      if (fillOverride) element.style.backgroundColor = fillOverride;
      return;
    }
    const { font, fill, align } = style;
    element.style.fontFamily = `"${font.family}", "Microsoft JhengHei", serif`;
    element.style.fontSize = `${Math.max(11, font.size * 1.333)}px`;
    element.style.fontWeight = font.bold ? '700' : '400';
    element.style.fontStyle = font.italic ? 'italic' : 'normal';
    element.style.textDecoration = [font.underline ? 'underline' : '', font.strike ? 'line-through' : ''].filter(Boolean).join(' ') || 'none';
    element.style.color = '#111111';
    if (fillOverride || fill) element.style.backgroundColor = fillOverride || makeFillReadableWithBlackText(fill);
    element.style.justifyContent = { center: 'center', right: 'flex-end', fill: 'stretch', justify: 'space-between' }[align.horizontal] || 'flex-start';
    element.style.alignItems = { center: 'center', top: 'flex-start', bottom: 'flex-end', justify: 'stretch' }[align.vertical] || 'flex-end';
    element.style.paddingLeft = `${4 + (align.indent || 0) * 10}px`;
    if (align.wrap) element.classList.add('wrap');
    if (align.rotation) element.style.writingMode = 'vertical-rl';
    element.style.borderTop = '0';
    element.style.borderLeft = '0';
    element.style.borderRight = '1px solid #dfe3e7';
    element.style.borderBottom = '1px solid #dfe3e7';
  }

  function selectionPosition(address) {
    return { row: rowNumber(address), column: columnNumber(columnLabel(address)) };
  }

  function updateSelection(element, address, content, options = {}) {
    const focus = selectionPosition(address);
    if (!options.extend && !options.preserveRange) selectionAnchor = focus;
    if (!selectionAnchor) selectionAnchor = focus;
    selectionFocus = focus;
    selectedCellElements.forEach((selected) => selected.classList.remove('selected', 'selection-focus'));
    selectedCellElements.clear();
    const startRow = Math.min(selectionAnchor.row, focus.row);
    const endRow = Math.max(selectionAnchor.row, focus.row);
    const startColumn = Math.min(selectionAnchor.column, focus.column);
    const endColumn = Math.max(selectionAnchor.column, focus.column);
    for (let row = startRow; row <= endRow; row += 1) {
      if (activeSheet.rows[row - 1]?.hidden || !isHistoryRowVisible(activeSheet, row)) continue;
      for (let column = startColumn; column <= endColumn; column += 1) {
        if (activeSheet.columns[column - 1]?.hidden) continue;
        const selected = renderedCellElements.get(cellKey(row, column));
        if (!selected) continue;
        selected.classList.add('selected');
        selectedCellElements.add(selected);
      }
    }
    activeCell = element || renderedCellElements.get(cellKey(focus.row, focus.column)) || null;
    if (activeCell) activeCell.classList.add('selected', 'selection-focus');
    activeSheet.activeCell = address;
    nameBox.textContent = address;
    formulaContent.textContent = content === undefined || content === null ? '' : String(content);
  }

  function nextVisibleSelection(sheet, start, rowStep, columnStep) {
    let row = start.row + rowStep;
    let column = start.column + columnStep;
    while (row >= 1 && row <= sheet.maxRow && column >= 1 && column <= sheet.maxCol) {
      if (!sheet.rows[row - 1]?.hidden && isHistoryRowVisible(sheet, row) && !sheet.columns[column - 1]?.hidden) {
        return { row, column };
      }
      row += rowStep;
      column += columnStep;
    }
    return null;
  }

  function handleSheetArrowKey(event) {
    if (event.target !== sheetScroll) return;
    const directions = {
      ArrowUp: [-1, 0],
      ArrowDown: [1, 0],
      ArrowLeft: [0, -1],
      ArrowRight: [0, 1],
    };
    const direction = directions[event.key];
    if (!direction) return;
    const start = selectionFocus || selectionPosition(activeSheet.activeCell || 'A1');
    const next = nextVisibleSelection(activeSheet, start, direction[0], direction[1]);
    if (!next) return;
    event.preventDefault();
    const address = `${activeSheet.columns[next.column - 1].letter}${next.row}`;
    const element = renderedCellElements.get(cellKey(next.row, next.column));
    updateSelection(element, address, getRaw(activeSheet, next.row, next.column), { extend: event.shiftKey });
    element?.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
  }

  function makeHeader(className, text, row, column) {
    const header = document.createElement('div');
    header.className = className;
    header.textContent = text;
    header.style.gridRow = row;
    header.style.gridColumn = column;
    return header;
  }

  function addColumnResizeHandle(header, sheet, index) {
    const handle = document.createElement('span');
    handle.className = 'column-resize-handle';
    handle.setAttribute('role', 'separator');
    handle.setAttribute('aria-label', `${sheet.columns[index].letter} 欄寬調整`);
    handle.setAttribute('aria-orientation', 'vertical');
    handle.addEventListener('pointerdown', (event) => {
      event.preventDefault();
      event.stopPropagation();
      const startX = event.clientX;
      const startWidth = sheet.columns[index].width;
      let finished = false;
      document.body.classList.add('resizing-column');
      const resize = (moveEvent) => {
        const width = Math.min(720, Math.max(34, Math.round(startWidth + moveEvent.clientX - startX)));
        if (width === sheet.columns[index].width) return;
        sheet.columns[index].width = width;
        renderGrid(sheet);
      };
      const finish = () => {
        if (finished) return;
        finished = true;
        document.body.classList.remove('resizing-column');
        document.removeEventListener('pointermove', resize);
        document.removeEventListener('pointerup', finish);
        document.removeEventListener('pointercancel', finish);
        saveState();
        syncHorizontalScrollControl();
      };
      document.addEventListener('pointermove', resize);
      document.addEventListener('pointerup', finish);
      document.addEventListener('pointercancel', finish);
    });
    header.append(handle);
  }

  function isEditable(sheet, row, column) {
    if (sheet.name === 'secret') return false;
    if (workbookUnlocked) {
      if (FORM_SHEETS.has(sheet.name) || sheet.name === '歷史資料') return row >= 3 && column <= sheet.maxCol;
      return row >= 2 && column <= sheet.maxCol;
    }
    if (FORM_SHEETS.has(sheet.name)) return row >= 3 && column >= 1 && column <= 13;
    if (sheet.name === '歷史資料') return row >= 3 && HISTORY_STANDARD_EDITABLE_COLUMNS.has(column);
    if (sheet.name === '設定') return row >= 2 && row <= 30 && column >= 1 && column <= 9;
    if (sheet.name === '統計') return (PAYMENT_ROWS.includes(row) || SUMMARY_ROWS.includes(row)) && (column === 5 || column === 7);
    return false;
  }

  function parseEditedValue(value, style) {
    const trimmed = value.trim();
    if (!trimmed) return '';
    const format = style?.numberFormat || '';
    const numeric = Number(trimmed.replace(/,/g, '').replace('%', ''));
    if (!Number.isFinite(numeric)) return trimmed;
    if (format.includes('%')) return trimmed.includes('%') ? numeric / 100 : numeric;
    return numeric;
  }

  function editCell(element, sheet, row, column, data) {
    if (element.querySelector('input')) return;
    const input = document.createElement('input');
    input.className = 'cell-editor';
    input.type = 'text';
    input.value = data?.raw ?? data?.v ?? '';
    input.setAttribute('aria-label', `${sheet.columns[column - 1].letter}${row}`);
    element.classList.add('editing');
    element.replaceChildren(input);
    input.focus();
    input.select();
    let committed = false;
    const commit = () => {
      if (committed) return;
      committed = true;
      const raw = parseEditedValue(input.value, sheet.styles[data?.s]);
      setCell(sheet, row, column, raw, { style: data?.s });
      // Excel keeps the same column selected and advances one visible row after an edit.
      let nextRow = row;
      for (let candidate = row + 1; candidate <= sheet.maxRow; candidate += 1) {
        if (!sheet.rows[candidate - 1]?.hidden) {
          nextRow = candidate;
          break;
        }
      }
      sheet.activeCell = `${sheet.columns[column - 1].letter}${nextRow}`;
      selectionAnchor = null;
      selectionFocus = null;
      statusNotice = '';
      if (FORM_SHEETS.has(sheet.name)) initializeFormRow(sheet, row);
      recalculate();
      saveState();
      renderGrid(sheet);
      renderTabs();
    };
    input.addEventListener('blur', commit);
    input.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') {
        event.preventDefault();
        input.blur();
      }
      if (event.key === 'Escape') {
        committed = true;
        renderGrid(sheet);
      }
    });
  }

  function formDataExists(sheet, row) {
    return Array.from({ length: 13 }, (_, index) => getRaw(sheet, row, index + 1)).some((value) => String(value ?? '').trim() !== '');
  }

  function amountInputs(sheet, row) {
    return amountColumns
      .map((column) => ({ column, raw: getRaw(sheet, row, column) }))
      .filter(({ raw }) => String(raw ?? '').trim() !== '');
  }

  function isCompleted(sheet, row) {
    return Boolean(getCell(sheet, row, 14)?.checked);
  }

  function setCompletion(sheet, row, checked) {
    setCell(sheet, row, 14, '', { checkbox: true, checked });
  }

  function initializeFormRow(sheet, row) {
    const key = getRaw(sheet, row, 16);
    if (!formDataExists(sheet, row)) {
      if (key) deleteHistoryByKey(key);
      [14, 15, 16].forEach((column) => removeCell(sheet, row, column));
      return;
    }
    if (!getRaw(sheet, row, 15)) setCell(sheet, row, 15, formatNow());
    if (!getRaw(sheet, row, 16)) {
      sequence += 1;
      setCell(sheet, row, 16, sequence);
      const secret = getSheet('secret');
      if (secret) setCell(secret, 1, 1, sequence);
    }
    if (!getCell(sheet, row, 14)?.checkbox) setCompletion(sheet, row, false);
    if (isCompleted(sheet, row)) syncFormToHistory(sheet, row);
  }

  function settingFor(platform) {
    const setting = getSheet('設定');
    if (!setting) return null;
    const row = Array.from({ length: setting.maxRow - 1 }, (_, index) => index + 2)
      .find((candidate) => String(getRaw(setting, candidate, 1)) === String(platform));
    if (!row) return null;
    return {
      platform: getRaw(setting, row, 1),
      role: getRaw(setting, row, 2),
      feePercent: numberValue(getRaw(setting, row, 4)),
      buyFee: numberValue(getRaw(setting, row, 5)),
      twd: numberValue(getRaw(setting, row, 6)),
      currencyRate: numberValue(getRaw(setting, row, 7)),
      matchRate: numberValue(getRaw(setting, row, 8)),
      bankFee: numberValue(getRaw(setting, row, 9)),
    };
  }

  function addOrUpdateHistory(record) {
    const history = getSheet('歷史資料');
    if (!history) return;
    let row = Array.from({ length: history.maxRow - 2 }, (_, index) => index + 3)
      .find((candidate) => String(getRaw(history, candidate, 1)) === String(record.key));
    if (!row) {
      row = Math.max(3, ...history.cells.filter((cell) => cell.r >= 3 && cell.v !== '').map((cell) => cell.r + 1));
      ensureRows(history, row);
    }
    const values = Object.fromEntries(Array.from({ length: 28 }, (_, index) => [index + 1, '']));
    Object.assign(values, {
      1: record.key, 2: record.openedAt, 3: formatNow(), 4: record.clientId, 5: record.platform,
      6: record.role, 7: record.nickname, 8: record.paymentMethod, 9: record.transactionType, 10: record.account,
    });
    if (record.transactionType === '收款') {
      const beforeFee = Math.trunc(record.amount * record.setting.currencyRate / record.setting.twd);
      const afterFee = beforeFee * (1 + record.setting.buyFee);
      const invoice = Math.trunc(record.amount * record.setting.feePercent);
      values[11] = record.amount;
      values[12] = record.amount;
      values[14] = invoice;
      values[18] = beforeFee;
      values[19] = record.setting.buyFee;
      values[20] = afterFee;
      values[27] = invoice;
    } else {
      const matchedCoin = Math.trunc(record.amount * record.setting.twd * record.setting.matchRate);
      const receivedCoin = Math.trunc(record.amount * record.setting.twd * record.setting.currencyRate);
      values[15] = record.amount;
      values[16] = record.amount;
      values[17] = record.setting.bankFee;
      values[21] = receivedCoin;
      values[22] = record.setting.matchRate;
      values[23] = matchedCoin;
      values[24] = receivedCoin + matchedCoin;
    }
    Object.entries(values).forEach(([column, value]) => setCell(history, row, Number(column), value));
  }

  function deleteHistoryByKey(key) {
    const history = getSheet('歷史資料');
    if (!history || key === '') return;
    const row = Array.from({ length: history.maxRow - 2 }, (_, index) => index + 3)
      .find((candidate) => String(getRaw(history, candidate, 1)) === String(key));
    if (!row) return;
    history.cells = history.cells
      .filter((cell) => cell.r !== row)
      .map((cell) => (cell.r > row ? { ...cell, r: cell.r - 1 } : cell));
  }

  function syncFormToHistory(sheet, row) {
    const amounts = amountInputs(sheet, row);
    const key = getRaw(sheet, row, 16);
    if (amounts.length !== 1 || !key) return false;
    const { column, raw } = amounts[0];
    const platform = getRaw(sheet, 2, column);
    const setting = settingFor(platform);
    if (!setting || !setting.twd) return false;
    addOrUpdateHistory({
      key,
      openedAt: getRaw(sheet, row, 15),
      clientId: getRaw(sheet, row, 1),
      nickname: getRaw(sheet, row, 2),
      account: getRaw(sheet, row, 3),
      platform,
      role: setting.role,
      paymentMethod: sheet.name,
      transactionType: column <= 8 ? '收款' : '出款',
      amount: numberValue(raw),
      setting,
    });
    return true;
  }

  function toggleCompletion(sheet, row, checked) {
    const amounts = amountInputs(sheet, row);
    if (checked && amounts.length !== 1) {
      setCompletion(sheet, row, false);
      statusNotice = '每筆交易只能填寫一個金額後結案。';
      recalculate();
      saveState();
      renderGrid(sheet);
      renderTabs();
      return;
    }
    if (checked && !syncFormToHistory(sheet, row)) {
      setCompletion(sheet, row, false);
      statusNotice = '找不到對應設定，或台幣換算值不可為 0。';
      recalculate();
      saveState();
      renderGrid(sheet);
      renderTabs();
      return;
    }
    if (!checked) deleteHistoryByKey(getRaw(sheet, row, 16));
    setCompletion(sheet, row, checked);
    statusNotice = '';
    recalculate();
    saveState();
    renderGrid(sheet);
    renderTabs();
  }

  function sumHistory(column, predicate = () => true) {
    const history = getSheet('歷史資料');
    if (!history) return 0;
    let total = 0;
    for (let row = 3; row <= history.maxRow; row += 1) {
      if (predicate(row)) total += numberValue(getRaw(history, row, column));
    }
    return total;
  }

  function updatePaymentDifferenceMarkers() {
    const history = getSheet('歷史資料');
    if (!history) return;
    for (let row = 3; row <= history.maxRow; row += 1) {
      const received = getRaw(history, row, 11);
      const actual = getRaw(history, row, 12);
      const receivedText = String(received ?? '').trim().replace(/,/g, '');
      const actualText = String(actual ?? '').trim().replace(/,/g, '');
      if (!receivedText || !actualText) {
        removeCell(history, row, 13);
        continue;
      }
      const receivedNumber = Number(receivedText);
      const actualNumber = Number(actualText);
      const differs = Number.isFinite(receivedNumber) && Number.isFinite(actualNumber)
        ? Math.abs(receivedNumber - actualNumber) > Number.EPSILON
        : receivedText !== actualText;
      if (differs) setCell(history, row, 13, '⚠');
      else removeCell(history, row, 13);
    }
  }

  function setDerived(sheet, row, column, raw) {
    setCell(sheet, row, column, raw);
  }

  function recalculateStatsRow(stats, row, predicate) {
    const receipts = sumHistory(12, predicate);
    const payouts = sumHistory(16, predicate);
    const bankFees = sumHistory(17, predicate);
    const receivedCoin = sumHistory(21, predicate);
    const matchedCoin = sumHistory(23, predicate);
    const goodsReceived = sumHistory(24, predicate);
    const goodsSent = sumHistory(20, predicate);
    setDerived(stats, row, 2, receipts);
    setDerived(stats, row, 3, payouts);
    setDerived(stats, row, 4, bankFees);
    setDerived(stats, row, 6, numberValue(getRaw(stats, row, 5)) + receipts - payouts - bankFees);
    setDerived(stats, row, 8, receivedCoin);
    setDerived(stats, row, 9, matchedCoin);
    setDerived(stats, row, 10, goodsReceived);
    setDerived(stats, row, 11, goodsSent);
    setDerived(stats, row, 12, goodsReceived + numberValue(getRaw(stats, row, 7)) - goodsSent);
  }

  function recalculateStats() {
    const stats = getSheet('統計');
    const settings = getSheet('設定');
    const history = getSheet('歷史資料');
    if (!stats || !settings || !history) return;
    PAYMENT_ROWS.forEach((row) => {
      const platform = getRaw(settings, row, 1);
      setDerived(stats, row, 1, platform);
      recalculateStatsRow(stats, row, (historyRow) => String(getRaw(history, historyRow, 5)) === String(platform));
    });
    SUMMARY_ROWS.forEach((row) => {
      const transaction = getRaw(stats, row, 1);
      recalculateStatsRow(stats, row, (historyRow) => String(getRaw(history, historyRow, 8)) === String(transaction));
    });
  }

  function recalculateFormHeaders() {
    const settings = getSheet('設定');
    if (!settings) return;
    workbook.sheets.filter((sheet) => FORM_SHEETS.has(sheet.name)).forEach((sheet) => {
      for (let offset = 0; offset < 5; offset += 1) {
        setDerived(sheet, 2, 4 + offset, getRaw(settings, 2 + offset, 1));
        setDerived(sheet, 2, 9 + offset, getRaw(settings, 2 + offset, 1));
      }
    });
  }

  function recalculate() {
    recalculateFormHeaders();
    updatePaymentDifferenceMarkers();
    recalculateStats();
  }

  function renderCheckbox(data, sheet, row) {
    const checkbox = document.createElement('input');
    checkbox.className = 'completion-checkbox';
    checkbox.type = 'checkbox';
    checkbox.checked = Boolean(data.checked);
    checkbox.setAttribute('aria-label', `${sheet.name} 第 ${row} 列結案`);
    checkbox.addEventListener('click', (event) => event.stopPropagation());
    checkbox.addEventListener('change', () => toggleCompletion(sheet, row, checkbox.checked));
    return checkbox;
  }

  function statusText(sheet) {
    if (statusNotice) return statusNotice;
    if (!workbookUnlocked) return '工作簿已上鎖 · 可修改原本開放的欄位；解鎖後才可編輯所有資料欄位。';
    if (FORM_SHEETS.has(sheet.name)) return `${sheet.name} · 已解鎖，雙擊資料欄位即可修改；勾選「結案」後自動寫入歷史與統計。`;
    if (sheet.name === '歷史資料') return '歷史資料 · 委託人編號、遊戲平台、經營角色與遊戲暱稱右下角的 ▼ 可篩選資料；收款與實收不同時會顯示 ⚠。';
    if (sheet.name === '設定') return '設定 · 已解鎖，雙擊資料欄位可修改；新結案交易會套用這裡的換算設定。';
    if (sheet.name === '統計') return '統計 · 已解鎖，雙擊資料欄位可修改；自動計算欄位會在重新計算時更新。';
    return `${sheet.name}${sheet.state === 'visible' ? '' : '（隱藏工作表）'} · 由交易與設定自動更新。`;
  }

  function renderGrid(sheet) {
    const cellMap = new Map(sheet.cells.map((cell) => [cellKey(cell.r, cell.c), cell]));
    const merges = mergeIndex(sheet);
    const frozen = freezePane(sheet);
    const rowHeaderWidth = 46;
    const columnHeaderHeight = 24;
    const displayColumns = displayColumnsFor(sheet, rowHeaderWidth);
    sheetCanvasWidth = Math.round(sheetScroll.clientWidth);
    const columnWidths = displayColumns.map((column) => (column.hidden ? 0 : column.width));
    const rowHeights = sheet.rows.map((row, index) => (row.hidden || !isHistoryRowVisible(sheet, index + 1) ? 0 : row.height));
    const frozenTop = [columnHeaderHeight];
    rowHeights.forEach((height, index) => { frozenTop[index + 1] = frozenTop[index] + height; });
    const frozenLeft = [rowHeaderWidth];
    columnWidths.forEach((width, index) => { frozenLeft[index + 1] = frozenLeft[index] + width; });

    grid.replaceChildren();
    renderedCellElements.clear();
    selectedCellElements.clear();
    activeCell = null;
    grid.style.gridTemplateColumns = [`${rowHeaderWidth}px`, ...columnWidths.map((width) => `${width}px`)].join(' ');
    grid.style.gridTemplateRows = [`${columnHeaderHeight}px`, ...rowHeights.map((height) => `${height}px`)].join(' ');
    grid.append(makeHeader('corner', '', 1, 1));

    displayColumns.forEach((column, index) => {
      if (column.hidden) return;
      const header = makeHeader('column-header', column.letter, 1, index + 2);
      if (!column.virtual) addColumnResizeHandle(header, sheet, index);
      if (index < frozen.columns) {
        header.classList.add('header-frozen');
        header.style.left = `${frozenLeft[index]}px`;
      }
      grid.append(header);
    });

    sheet.rows.forEach((row, index) => {
      if (row.hidden || !isHistoryRowVisible(sheet, index + 1)) return;
      const header = makeHeader('row-header', String(index + 1), index + 2, 1);
      if (index < frozen.rows) {
        header.classList.add('header-frozen');
        header.style.top = `${frozenTop[index]}px`;
      }
      grid.append(header);
    });

    for (let row = 1; row <= sheet.maxRow; row += 1) {
      if (sheet.rows[row - 1]?.hidden || !isHistoryRowVisible(sheet, row)) continue;
      for (let column = 1; column <= displayColumns.length; column += 1) {
        const columnDefinition = displayColumns[column - 1];
        if (columnDefinition.hidden) continue;
        const isVirtualColumn = Boolean(columnDefinition.virtual);
        const merge = isVirtualColumn ? null : merges.get(cellKey(row, column));
        if (merge && (merge.top !== row || merge.left !== column)) continue;
        const data = cellMap.get(cellKey(row, column));
        const cell = document.createElement('div');
        const colorOverride = isVirtualColumn ? '' : sectionFill(sheet, row, column);
        cell.className = 'grid-cell';
        cell.style.gridRow = `${row + 1} / span ${merge?.rowSpan || 1}`;
        cell.style.gridColumn = `${column + 1} / span ${merge?.columnSpan || 1}`;
        if (isVirtualColumn) {
          cell.classList.add('canvas-spacer');
          grid.append(cell);
          continue;
        }
        if (data?.v || data?.checkbox) cell.classList.add('has-value');
        if (data) applyStyle(cell, sheet.styles[data.s], colorOverride);
        else if (colorOverride) cell.style.backgroundColor = colorOverride;
        if (sheet.name === '歷史資料' && row >= 3 && column === 13 && data?.v === '⚠') {
          cell.classList.add('difference-marker');
          cell.title = '收款與實收金額不同';
        }
        const address = `${sheet.columns[column - 1].letter}${row}`;
        if (data?.checkbox) cell.append(renderCheckbox(data, sheet, row));
        else cell.textContent = data?.v || '';
        const filterButton = renderHistoryFilterButton(sheet, row, column);
        if (filterButton) cell.append(filterButton);
        if (isEditable(sheet, row, column)) {
          cell.classList.add('editable');
          cell.title = '雙擊修改';
          cell.addEventListener('dblclick', () => editCell(cell, sheet, row, column, data));
        } else if (!workbookUnlocked && row >= 2 && sheet.name !== 'secret') {
          cell.classList.add('locked');
          cell.title = '工作簿已上鎖';
        }
        const isFrozenRow = row <= frozen.rows;
        const isFrozenColumn = column <= frozen.columns;
        if (isFrozenRow || isFrozenColumn) {
          cell.classList.add('frozen');
          if (isFrozenRow) cell.style.top = `${frozenTop[row - 1]}px`;
          if (isFrozenColumn) cell.style.left = `${frozenLeft[column - 1]}px`;
        }
        cell.addEventListener('click', (event) => updateSelection(cell, address, data?.raw ?? data?.v, { extend: event.shiftKey }));
        for (let mergedRow = merge?.top || row; mergedRow <= (merge?.top || row) + (merge?.rowSpan || 1) - 1; mergedRow += 1) {
          for (let mergedColumn = merge?.left || column; mergedColumn <= (merge?.left || column) + (merge?.columnSpan || 1) - 1; mergedColumn += 1) {
            renderedCellElements.set(cellKey(mergedRow, mergedColumn), cell);
          }
        }
        grid.append(cell);
        if (address === sheet.activeCell) updateSelection(cell, address, data?.raw ?? data?.v, { preserveRange: true });
      }
    }
    statusBar.textContent = statusText(sheet);
    requestAnimationFrame(syncHorizontalScrollControl);
  }

  function renderTabs() {
    tabs.replaceChildren();
    workbook.sheets
      .filter((sheet) => (sheet.state === 'visible' || (sheet.name === '設定' && workbookUnlocked)) && canOpenSheet(sheet))
      .forEach((sheet) => {
        const tab = document.createElement('button');
        tab.type = 'button';
        tab.className = `sheet-tab${sheet === activeSheet ? ' active' : ''}`;
        tab.textContent = sheet.name;
        tab.addEventListener('click', () => setSheet(sheet));
        tabs.append(tab);
      });
    const hiddenSheets = workbook.sheets.filter((sheet) => sheet.state === 'hidden' && sheet.name !== '設定');
    hiddenButton.hidden = hiddenSheets.length === 0;
    hiddenButton.textContent = `隱藏工作表 (${hiddenSheets.length})`;
    hiddenMenu.replaceChildren();
    hiddenSheets.forEach((sheet) => {
      const item = document.createElement('button');
      item.className = 'hidden-sheet-item';
      item.type = 'button';
      item.textContent = sheet.name;
      item.addEventListener('click', () => {
        hiddenMenu.hidden = true;
        hiddenButton.setAttribute('aria-expanded', 'false');
        setSheet(sheet);
      });
      hiddenMenu.append(item);
    });
  }

  function setSheet(sheet) {
    if (!canOpenSheet(sheet)) return;
    activeSheet = sheet;
    activeCell = null;
    selectionAnchor = null;
    selectionFocus = null;
    statusNotice = '';
    sheetScroll.scrollTo({ top: 0, left: 0, behavior: 'auto' });
    closeHistoryFilterMenu();
    renderGrid(sheet);
    renderTabs();
  }

  hiddenButton.addEventListener('click', () => {
    hiddenMenu.hidden = !hiddenMenu.hidden;
    hiddenButton.setAttribute('aria-expanded', String(!hiddenMenu.hidden));
  });
  horizontalScrollLeft.addEventListener('click', () => {
    moveHorizontalScroll(sheetScroll.scrollLeft - Math.max(48, sheetScroll.clientWidth * 0.2));
  });
  horizontalScrollRight.addEventListener('click', () => {
    moveHorizontalScroll(sheetScroll.scrollLeft + Math.max(48, sheetScroll.clientWidth * 0.2));
  });
  horizontalScrollTrack.addEventListener('pointerdown', (event) => {
    if (event.target === horizontalScrollThumb) return;
    event.preventDefault();
    scrollFromTrackPosition(event.clientX);
  });
  horizontalScrollThumb.addEventListener('pointerdown', (event) => {
    event.preventDefault();
    event.stopPropagation();
    const startX = event.clientX;
    const startScroll = sheetScroll.scrollLeft;
    let finished = false;
    horizontalScrollThumb.classList.add('dragging');
    const drag = (moveEvent) => {
      const maximum = Math.max(0, sheetScroll.scrollWidth - sheetScroll.clientWidth);
      const trackWidth = horizontalScrollTrack.getBoundingClientRect().width;
      const thumbWidth = horizontalScrollThumb.getBoundingClientRect().width;
      const travel = Math.max(0, trackWidth - thumbWidth);
      moveHorizontalScroll(travel ? startScroll + maximum * (moveEvent.clientX - startX) / travel : 0);
    };
    const finish = () => {
      if (finished) return;
      finished = true;
      horizontalScrollThumb.classList.remove('dragging');
      document.removeEventListener('pointermove', drag);
      document.removeEventListener('pointerup', finish);
      document.removeEventListener('pointercancel', finish);
    };
    document.addEventListener('pointermove', drag);
    document.addEventListener('pointerup', finish);
    document.addEventListener('pointercancel', finish);
  });
  horizontalScrollTrack.addEventListener('keydown', (event) => {
    const maximum = Math.max(0, sheetScroll.scrollWidth - sheetScroll.clientWidth);
    const step = Math.max(48, sheetScroll.clientWidth * 0.1);
    const actions = {
      ArrowLeft: () => moveHorizontalScroll(sheetScroll.scrollLeft - step),
      ArrowRight: () => moveHorizontalScroll(sheetScroll.scrollLeft + step),
      PageUp: () => moveHorizontalScroll(sheetScroll.scrollLeft - sheetScroll.clientWidth),
      PageDown: () => moveHorizontalScroll(sheetScroll.scrollLeft + sheetScroll.clientWidth),
      Home: () => moveHorizontalScroll(0),
      End: () => moveHorizontalScroll(maximum),
    };
    if (!actions[event.key]) return;
    event.preventDefault();
    actions[event.key]();
  });
  sheetScroll.addEventListener('keydown', handleSheetArrowKey);
  sheetScroll.addEventListener('scroll', () => {
    syncHorizontalScrollControl();
    closeHistoryFilterMenu();
  }, { passive: true });
  if ('ResizeObserver' in window) {
    const resizeObserver = new ResizeObserver(() => requestAnimationFrame(() => {
      syncHorizontalScrollControl();
      const width = Math.round(sheetScroll.clientWidth);
      if (width === sheetCanvasWidth) return;
      renderGrid(activeSheet);
    }));
    resizeObserver.observe(sheetScroll);
    resizeObserver.observe(horizontalScrollTrack);
  }
  document.addEventListener('click', (event) => {
    if (!event.target.closest('.hidden-sheet-wrap')) {
      hiddenMenu.hidden = true;
      hiddenButton.setAttribute('aria-expanded', 'false');
    }
    if (!event.target.closest('.history-filter-menu') && !event.target.closest('.column-filter-button')) closeHistoryFilterMenu();
  });

  const lockReady = initializeWorkbookLock();
  updateLockControls();
  workbookLockButton.addEventListener('click', () => {
    if (workbookUnlocked) {
      workbookUnlocked = false;
      statusNotice = '工作簿已上鎖。';
      if (!canOpenSheet(activeSheet)) {
        activeSheet = workbook.sheets.find((sheet) => sheet.state === 'visible' && canOpenSheet(sheet)) || workbook.sheets[0];
        activeCell = null;
        selectionAnchor = null;
        selectionFocus = null;
        sheetScroll.scrollTo({ top: 0, left: 0, behavior: 'auto' });
      }
      updateLockControls();
      renderGrid(activeSheet);
      renderTabs();
      return;
    }
    openLockDialog(lockConfig.passwordHash ? 'unlock' : 'setup-password');
  });
  changePasswordButton.addEventListener('click', () => {
    if (workbookUnlocked) openLockDialog('change-password');
  });
  lockDialogCancel.addEventListener('click', () => lockDialog.close());
  lockDialogForm.addEventListener('submit', handleLockDialogSubmit);

  const secret = getSheet('secret');
  revealInvoiceAmountColumn();
  sequence = Math.trunc(numberValue(secret ? getRaw(secret, 1, 1) : 0));
  ensureWorkbookCanvas();
  workbook.sheets.filter((sheet) => FORM_SHEETS.has(sheet.name)).forEach((sheet) => {
    for (let row = 3; row <= sheet.maxRow; row += 1) initializeFormRow(sheet, row);
  });
  recalculate();
  setSheet(activeSheet);
})();
