(() => {
  const workbook = window.GRIDBOOK;
  const FORM_SHEETS = new Set(['銀行轉帳', '超商轉帳', '虛擬帳號', '其他入帳']);
  const PAYMENT_ROWS = [2, 3, 4, 5, 6];
  const SUMMARY_ROWS = [13, 14, 15, 16, 17];
  const amountColumns = Array.from({ length: 10 }, (_, index) => index + 4);
  const HISTORY_EDITABLE_COLUMNS = new Set([12, 13, 15, 16, 24, 26]);
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
  const storageKey = 'v20-12-editable-workbook-v1';

  workbook.sheets = workbook.sheets.filter((sheet) => sheet.name !== '發票');
  try {
    const saved = JSON.parse(localStorage.getItem(storageKey) || 'null');
    if (saved?.sheets?.length) workbook.sheets = saved.sheets.filter((sheet) => sheet.name !== '發票');
  } catch {
    localStorage.removeItem(storageKey);
  }

  let activeSheet = workbook.sheets.find((sheet) => sheet.state === 'visible') || workbook.sheets[0];
  let activeCell = null;
  let sequence = 0;
  let statusNotice = '';

  const cellKey = (row, column) => `${row}:${column}`;
  const columnNumber = (label) => {
    let value = 0;
    for (const character of label) value = value * 26 + character.charCodeAt(0) - 64;
    return value;
  };
  const rowNumber = (label) => Number(label.match(/\d+/)?.[0] || 1);
  const columnLabel = (label) => label.match(/[A-Z]+/)?.[0] || 'A';
  const getSheet = (name) => workbook.sheets.find((sheet) => sheet.name === name);

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
    const invoiceAmountColumn = history?.columns?.[25];
    if (invoiceAmountColumn) invoiceAmountColumn.hidden = false;
  }

  function getCell(sheet, row, column) {
    return sheet.cells.find((cell) => cell.r === row && cell.c === column);
  }

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

  function formatNow() {
    const now = new Date();
    const pad = (value) => String(value).padStart(2, '0');
    return `${now.getFullYear()}/${pad(now.getMonth() + 1)}/${pad(now.getDate())} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
  }

  function saveState() {
    localStorage.setItem(storageKey, JSON.stringify({ sheets: workbook.sheets }));
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

  function updateSelection(element, address, content) {
    if (activeCell) activeCell.classList.remove('selected');
    activeCell = element;
    if (element) element.classList.add('selected');
    nameBox.textContent = address;
    formulaContent.textContent = content === undefined || content === null ? '' : String(content);
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
    if (FORM_SHEETS.has(sheet.name)) return row >= 3 && column >= 1 && column <= 13;
    if (sheet.name === '歷史資料') return row >= 3 && HISTORY_EDITABLE_COLUMNS.has(column);
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
    const values = {
      1: record.key, 2: record.openedAt, 3: formatNow(), 4: record.clientId, 5: record.platform,
      6: record.role, 7: record.nickname, 8: record.paymentMethod, 9: record.transactionType, 10: record.account,
      11: '', 12: '', 13: '', 14: '', 15: '', 16: '', 17: '', 18: '', 19: '', 20: '', 21: '', 22: '', 23: '', 24: '', 25: '', 26: '', 27: '',
    };
    if (record.transactionType === '收款') {
      const beforeFee = Math.trunc(record.amount * record.setting.currencyRate / record.setting.twd);
      const afterFee = beforeFee * (1 + record.setting.buyFee);
      const invoice = Math.trunc(record.amount * record.setting.feePercent);
      values[11] = record.amount;
      values[12] = record.amount;
      values[13] = invoice;
      values[17] = beforeFee;
      values[18] = record.setting.buyFee;
      values[19] = afterFee;
      values[26] = invoice;
    } else {
      const matchedCoin = Math.trunc(record.amount * record.setting.twd * record.setting.matchRate);
      const receivedCoin = Math.trunc(record.amount * record.setting.twd * record.setting.currencyRate);
      values[14] = record.amount;
      values[15] = record.amount;
      values[16] = record.setting.bankFee;
      values[20] = receivedCoin;
      values[21] = record.setting.matchRate;
      values[22] = matchedCoin;
      values[23] = receivedCoin + matchedCoin;
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

  function setDerived(sheet, row, column, raw) {
    setCell(sheet, row, column, raw);
  }

  function recalculateHistoryTotals() {
    const history = getSheet('歷史資料');
    if (!history) return;
    const receipts = sumHistory(11);
    const payouts = sumHistory(14) + sumHistory(16);
    setDerived(history, 2, 28, receipts);
    setDerived(history, 2, 29, payouts);
    setDerived(history, 2, 31, numberValue(getRaw(history, 2, 30)) + receipts - payouts);
    setDerived(history, 2, 32, sumHistory(20));
    setDerived(history, 2, 33, sumHistory(22));
    setDerived(history, 2, 34, sumHistory(23));
    setDerived(history, 2, 35, sumHistory(19));
  }

  function recalculateStatsRow(stats, row, predicate) {
    const receipts = sumHistory(12, predicate);
    const payouts = sumHistory(15, predicate);
    const bankFees = sumHistory(16, predicate);
    const receivedCoin = sumHistory(20, predicate);
    const matchedCoin = sumHistory(22, predicate);
    const goodsReceived = sumHistory(23, predicate);
    const goodsSent = sumHistory(19, predicate);
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
    recalculateHistoryTotals();
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
    if (FORM_SHEETS.has(sheet.name)) return `${sheet.name} · 雙擊 A–M 欄輸入資料，勾選「結案」後自動寫入歷史與統計。`;
    if (sheet.name === '歷史資料') return '歷史資料 · 雙擊「實收、交易服務費、實支、手續費、備註、發票金額」可修正；統計會立即更新。';
    if (sheet.name === '設定') return '設定 · 雙擊可修改；新結案交易會套用這裡的換算設定。';
    if (sheet.name === '統計') return '統計 · E 與 G 欄可調整初始台幣與初始貨幣，其他欄位自動計算。';
    return `${sheet.name}${sheet.state === 'visible' ? '' : '（隱藏工作表）'} · 由交易與設定自動更新。`;
  }

  function renderGrid(sheet) {
    const cellMap = new Map(sheet.cells.map((cell) => [cellKey(cell.r, cell.c), cell]));
    const merges = mergeIndex(sheet);
    const frozen = freezePane(sheet);
    const rowHeaderWidth = 46;
    const columnHeaderHeight = 24;
    const columnWidths = sheet.columns.map((column) => (column.hidden ? 0 : column.width));
    const rowHeights = sheet.rows.map((row) => (row.hidden ? 0 : row.height));
    const frozenTop = [columnHeaderHeight];
    rowHeights.forEach((height, index) => { frozenTop[index + 1] = frozenTop[index] + height; });
    const frozenLeft = [rowHeaderWidth];
    columnWidths.forEach((width, index) => { frozenLeft[index + 1] = frozenLeft[index] + width; });

    grid.replaceChildren();
    grid.style.gridTemplateColumns = [`${rowHeaderWidth}px`, ...columnWidths.map((width) => `${width}px`)].join(' ');
    grid.style.gridTemplateRows = [`${columnHeaderHeight}px`, ...rowHeights.map((height) => `${height}px`)].join(' ');
    grid.append(makeHeader('corner', '', 1, 1));

    sheet.columns.forEach((column, index) => {
      if (column.hidden) return;
      const header = makeHeader('column-header', column.letter, 1, index + 2);
      addColumnResizeHandle(header, sheet, index);
      if (index < frozen.columns) {
        header.classList.add('header-frozen');
        header.style.left = `${frozenLeft[index]}px`;
      }
      grid.append(header);
    });

    sheet.rows.forEach((row, index) => {
      if (row.hidden) return;
      const header = makeHeader('row-header', String(index + 1), index + 2, 1);
      if (index < frozen.rows) {
        header.classList.add('header-frozen');
        header.style.top = `${frozenTop[index]}px`;
      }
      grid.append(header);
    });

    for (let row = 1; row <= sheet.maxRow; row += 1) {
      if (sheet.rows[row - 1]?.hidden) continue;
      for (let column = 1; column <= sheet.maxCol; column += 1) {
        if (sheet.columns[column - 1]?.hidden) continue;
        const merge = merges.get(cellKey(row, column));
        if (merge && (merge.top !== row || merge.left !== column)) continue;
        const data = cellMap.get(cellKey(row, column));
        const cell = document.createElement('div');
        const colorOverride = sectionFill(sheet, row, column);
        cell.className = 'grid-cell';
        cell.style.gridRow = `${row + 1} / span ${merge?.rowSpan || 1}`;
        cell.style.gridColumn = `${column + 1} / span ${merge?.columnSpan || 1}`;
        if (data?.v || data?.checkbox) cell.classList.add('has-value');
        if (data) applyStyle(cell, sheet.styles[data.s], colorOverride);
        else if (colorOverride) cell.style.backgroundColor = colorOverride;
        const address = `${sheet.columns[column - 1].letter}${row}`;
        if (data?.checkbox) cell.append(renderCheckbox(data, sheet, row));
        else cell.textContent = data?.v || '';
        if (isEditable(sheet, row, column)) {
          cell.classList.add('editable');
          cell.title = '雙擊修改';
          cell.addEventListener('dblclick', () => editCell(cell, sheet, row, column, data));
        }
        const isFrozenRow = row <= frozen.rows;
        const isFrozenColumn = column <= frozen.columns;
        if (isFrozenRow || isFrozenColumn) {
          cell.classList.add('frozen');
          if (isFrozenRow) cell.style.top = `${frozenTop[row - 1]}px`;
          if (isFrozenColumn) cell.style.left = `${frozenLeft[column - 1]}px`;
        }
        cell.addEventListener('click', () => updateSelection(cell, address, data?.raw ?? data?.v));
        grid.append(cell);
        if (address === sheet.activeCell) updateSelection(cell, address, data?.raw ?? data?.v);
      }
    }
    statusBar.textContent = statusText(sheet);
    requestAnimationFrame(syncHorizontalScrollControl);
  }

  function renderTabs() {
    tabs.replaceChildren();
    workbook.sheets.filter((sheet) => sheet.state === 'visible').forEach((sheet) => {
      const tab = document.createElement('button');
      tab.type = 'button';
      tab.className = `sheet-tab${sheet === activeSheet ? ' active' : ''}`;
      tab.textContent = sheet.name;
      tab.addEventListener('click', () => setSheet(sheet));
      tabs.append(tab);
    });
    const hiddenSheets = workbook.sheets.filter((sheet) => sheet.state === 'hidden');
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
    activeSheet = sheet;
    activeCell = null;
    statusNotice = '';
    sheetScroll.scrollTo({ top: 0, left: 0, behavior: 'auto' });
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
  sheetScroll.addEventListener('scroll', syncHorizontalScrollControl, { passive: true });
  if ('ResizeObserver' in window) {
    const resizeObserver = new ResizeObserver(() => requestAnimationFrame(syncHorizontalScrollControl));
    resizeObserver.observe(sheetScroll);
    resizeObserver.observe(horizontalScrollTrack);
  }
  document.addEventListener('click', (event) => {
    if (!event.target.closest('.hidden-sheet-wrap')) {
      hiddenMenu.hidden = true;
      hiddenButton.setAttribute('aria-expanded', 'false');
    }
  });

  const secret = getSheet('secret');
  revealInvoiceAmountColumn();
  sequence = Math.trunc(numberValue(secret ? getRaw(secret, 1, 1) : 0));
  workbook.sheets.filter((sheet) => FORM_SHEETS.has(sheet.name)).forEach((sheet) => {
    for (let row = 3; row <= sheet.maxRow; row += 1) initializeFormRow(sheet, row);
  });
  recalculate();
  setSheet(activeSheet);
})();
