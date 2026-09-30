import ExcelJS from 'exceljs';
import FileSaver from 'file-saver';
import { PlaceholderVariable, TemplateMeta, CellOccurrence, QuotationItem } from '../types/template';

// Universal placeholder regex matching {{ variable_name | default_value }} or {{ variable_name }}
export const PLACEHOLDER_REGEX = /\{\{\s*([^\|\}]+?)(?:\s*\|\s*([^\{\}]+?))?\s*\}\}/g;

/**
 * Common English to Traditional Chinese placeholder key translation dictionary
 */
export const KEY_TRANSLATION_MAP: Record<string, string> = {
  COMPANY_LOGO: '公司商標名稱',
  QUOTATION_NO: '報價單號',
  QUOTATION_DATE: '報價日期',
  VALID_DAYS: '有效天數',
  PROJECT_NAME: '專案名稱',
  PROVIDER_NAME: '報價公司名稱',
  CLIENT_NAME: '客戶公司名稱',
  PROVIDER_TAX_ID: '報價公司統編',
  CLIENT_TAX_ID: '客戶統編',
  PROVIDER_PHONE: '報價公司電話',
  CLIENT_CONTACT: '客戶聯絡窗口',
  PROVIDER_ADDRESS: '報價公司地址',
  CLIENT_ADDRESS: '客戶通訊地址',
  客戶名稱: '客戶公司名稱',
  買受機構名稱: '客戶公司名稱',
  買受機構統編: '客戶統編',
  ITEM_1_NAME: '品項1_名稱',
  ITEM_1_DESC: '品項1_說明',
  ITEM_1_QTY: '品項1_數量',
  ITEM_1_UNIT: '品項1_單位',
  ITEM_1_PRICE: '品項1_單價',
  ITEM_2_NAME: '品項2_名稱',
  ITEM_2_DESC: '品項2_說明',
  ITEM_2_QTY: '品項2_數量',
  ITEM_2_UNIT: '品項2_單位',
  ITEM_2_PRICE: '品項2_單價',
  ITEM_3_NAME: '品項3_名稱',
  ITEM_3_DESC: '品項3_說明',
  ITEM_3_QTY: '品項3_數量',
  ITEM_3_UNIT: '品項3_單位',
  ITEM_3_PRICE: '品項3_單價',
  ITEM_4_NAME: '品項4_名稱',
  ITEM_4_DESC: '品項4_說明',
  ITEM_4_QTY: '品項4_數量',
  ITEM_4_UNIT: '品項4_單位',
  ITEM_4_PRICE: '品項4_單價',
  ITEM_5_NAME: '品項5_名稱',
  ITEM_5_DESC: '品項5_說明',
  ITEM_5_QTY: '品項5_數量',
  ITEM_5_UNIT: '品項5_單位',
  ITEM_5_PRICE: '品項5_單價',
  PAYMENT_TERMS: '付款條件說明',
  BANK_ACCOUNT: '匯款專戶資訊',
  VALID_TERMS: '有效期限條款',

  INVOICE_NO: '請款單號',
  INVOICE_DATE: '請款日期',
  DUE_DATE: '付款截止日',
  CONTRACT_NAME: '專案合約名稱',
  VENDOR_NAME: '受款公司名稱',
  CLIENT_COMPANY: '客戶公司名稱',
  VENDOR_TAX_ID: '受款公司統編',
  VENDOR_PHONE: '受款公司電話',
  VENDOR_ADDRESS: '受款公司通訊地址',
  STAGE_1_NAME: '請款階段1_項目名稱',
  STAGE_1_DESC: '請款階段1_工作內容說明',
  STAGE_1_QTY: '請款階段1_數量',
  STAGE_1_UNIT: '請款階段1_單位',
  STAGE_1_PRICE: '請款階段1_單價',
  STAGE_2_NAME: '請款階段2_項目名稱',
  STAGE_2_DESC: '請款階段2_工作內容說明',
  STAGE_2_QTY: '請款階段2_數量',
  STAGE_2_UNIT: '請款階段2_單位',
  STAGE_2_PRICE: '請款階段2_單價',
  STAGE_3_NAME: '請款階段3_項目名稱',
  STAGE_3_DESC: '請款階段3_工作內容說明',
  STAGE_3_QTY: '請款階段3_數量',
  STAGE_3_UNIT: '請款階段3_單位',
  STAGE_3_PRICE: '請款階段3_單價',
  BANK_NAME: '受款銀行名稱',
  BANK_BRANCH: '受款銀行分行',
  BANK_ACCOUNT_NAME: '受款戶名',
  PAYMENT_NOTE_1: '款項備註說明_1',
  PAYMENT_NOTE_2: '款項備註說明_2',
};

/**
 * Format any placeholder key into natural Traditional Chinese label
 */
export function formatKeyToChinese(key: string): string {
  if (KEY_TRANSLATION_MAP[key]) return KEY_TRANSLATION_MAP[key];
  if (KEY_TRANSLATION_MAP[key.toUpperCase()]) return KEY_TRANSLATION_MAP[key.toUpperCase()];
  return key;
}

/**
 * Categorize variable based on its key name (bilingual: English and Traditional Chinese)
 */
export function categorizeKey(key: string): string {
  const k = key.toUpperCase();
  if (
    k.includes('COMPANY') ||
    k.includes('PROVIDER') ||
    k.includes('CLIENT') ||
    k.includes('VENDOR') ||
    k.includes('TAX_ID') ||
    k.includes('PHONE') ||
    k.includes('CONTACT') ||
    k.includes('ADDRESS') ||
    k.includes('LOGO') ||
    k.includes('公司') ||
    k.includes('客戶') ||
    k.includes('人員') ||
    k.includes('聯絡') ||
    k.includes('電話') ||
    k.includes('信箱') ||
    k.includes('郵件') ||
    k.includes('廠商') ||
    k.includes('部門') ||
    k.includes('地址')
  ) {
    return '基本資訊';
  }
  if (
    k.includes('QUOTATION') ||
    k.includes('INVOICE') ||
    k.includes('NO') ||
    k.includes('DATE') ||
    k.includes('VALID') ||
    k.includes('DUE') ||
    k.includes('PROJECT') ||
    k.includes('CONTRACT') ||
    k.includes('單號') ||
    k.includes('編號') ||
    k.includes('日期') ||
    k.includes('專案') ||
    k.includes('有效') ||
    k.includes('付款日')
  ) {
    return '單據資訊';
  }
  if (
    k.includes('ITEM') ||
    k.includes('STAGE') ||
    k.includes('DESC') ||
    k.includes('UNIT') ||
    k.includes('QTY') ||
    k.includes('PRICE') ||
    k.includes('AMOUNT') ||
    k.includes('品項') ||
    k.includes('項目') ||
    k.includes('規格') ||
    k.includes('數量') ||
    k.includes('單價') ||
    k.includes('金額') ||
    k.includes('費用') ||
    k.includes('科目')
  ) {
    return '明細項目';
  }
  if (
    k.includes('PAYMENT') ||
    k.includes('TERMS') ||
    k.includes('NOTE') ||
    k.includes('MEMO') ||
    k.includes('條款') ||
    k.includes('備註') ||
    k.includes('說明') ||
    k.includes('備忘') ||
    k.includes('協議') ||
    k.includes('注意')
  ) {
    return '條款與備註';
  }
  if (
    k.includes('BANK') ||
    k.includes('BRANCH') ||
    k.includes('ACCOUNT') ||
    k.includes('銀行') ||
    k.includes('帳號') ||
    k.includes('戶名') ||
    k.includes('匯款') ||
    k.includes('支票')
  ) {
    return '帳戶金融';
  }
  return '其他欄位';
}

/**
 * Determine if a variable is a date variable
 */
export function isDateKey(key: string, defaultValue: string): boolean {
  const upper = key.toUpperCase();
  if (
    upper.includes('DATE') ||
    upper.includes('日期') ||
    upper.includes('截止日') ||
    upper.includes('付款日') ||
    key.includes('報價日期') ||
    key.includes('請款日期')
  ) {
    return true;
  }
  if (/^\d{4}[-/]\d{1,2}[-/]\d{1,2}$/.test(defaultValue.trim())) {
    return true;
  }
  return false;
}

/**
 * Determine if a variable should use Textarea
 */
export function isLongTextKey(key: string, defaultValue: string): boolean {
  if (defaultValue.includes('\n') || defaultValue.length > 30) return true;
  const upper = key.toUpperCase();
  const longKeywords = [
    'DESC',
    'TERMS',
    'NOTE',
    'MEMO',
    'ADDRESS',
    'PROJECT_NAME',
    'CONTRACT_NAME',
    '條款',
    '備註',
    '說明',
    '備忘',
    '協議',
    '內容',
    '地址',
    '清單',
    '事項',
  ];
  return longKeywords.some((w) => upper.includes(w) || key.includes(w));
}

/**
 * Predefined options dictionary for common business fields
 */
export const DROPDOWN_OPTIONS_MAP: Record<string, string[]> = {
  // 有效天數
  '有效天數': ['7 天', '14 天', '30 天', '45 天', '60 天', '90 天', '180 天', '一年內有效'],
  'VALID_DAYS': ['7 天', '14 天', '30 天', '45 天', '60 天', '90 天', '180 天', '一年內有效'],

  // 品項與請款單位
  '品項1_單位': ['專案', '式', '套件', '期', '場次', '個月', '小時', '人天', '個', '組', '台', '份'],
  '品項2_單位': ['專案', '式', '套件', '期', '場次', '個月', '小時', '人天', '個', '組', '台', '份'],
  '品項3_單位': ['專案', '式', '套件', '期', '場次', '個月', '小時', '人天', '個', '組', '台', '份'],
  '品項4_單位': ['專案', '式', '套件', '期', '場次', '個月', '小時', '人天', '個', '組', '台', '份'],
  '品項5_單位': ['專案', '式', '套件', '期', '場次', '個月', '小時', '人天', '個', '組', '台', '份'],
  '請款階段1_單位': ['期', '式', '專案', '套件', '個月', '小時', '人天', '場次', '份'],
  '請款階段2_單位': ['期', '式', '專案', '套件', '個月', '小時', '人天', '場次', '份'],
  '請款階段3_單位': ['期', '式', '專案', '套件', '個月', '小時', '人天', '場次', '份'],

  // 付款條件
  '付款條件說明': [
    '簽約訂金 30%、期中展示驗收 40%、上線結案驗收 30%',
    '簽約訂金 50%、完工驗收 50%',
    '驗收合格後次月 15 日電匯撥款',
    '驗收合格後 30 天電匯付款',
    '簽約全額預付 100%',
  ],
  'PAYMENT_TERMS': [
    '簽約訂金 30%、期中展示驗收 40%、上線結案驗收 30%',
    '簽約訂金 50%、完工驗收 50%',
    '驗收合格後次月 15 日電匯撥款',
    '驗收合格後 30 天電匯付款',
    '簽約全額預付 100%',
  ],

  // 銀行
  '受款銀行名稱': [
    '國泰世華商業銀行 (013)',
    '玉山商業銀行 (808)',
    '中國信託商業銀行 (822)',
    '台北富邦商業銀行 (012)',
    '台新國際商業銀行 (812)',
    '第一商業銀行 (007)',
    '合作金庫商業銀行 (006)',
    '兆豐國際商業銀行 (017)',
    '華南商業銀行 (008)',
    '永豐商業銀行 (807)',
  ],
  'BANK_NAME': [
    '國泰世華商業銀行 (013)',
    '玉山商業銀行 (808)',
    '中國信託商業銀行 (822)',
    '台北富邦商業銀行 (012)',
    '台新國際商業銀行 (812)',
    '第一商業銀行 (007)',
    '合作金庫商業銀行 (006)',
    '兆豐國際商業銀行 (017)',
    '華南商業銀行 (008)',
    '永豐商業銀行 (807)',
  ],
};

/**
 * Determine if a variable should have dropdown options
 */
export function getDropdownOptions(key: string, defaultValue?: string): string[] | null {
  const direct = DROPDOWN_OPTIONS_MAP[key] || DROPDOWN_OPTIONS_MAP[key.toUpperCase()];
  if (direct) {
    if (defaultValue && !direct.includes(defaultValue.trim()) && defaultValue.trim() !== '') {
      return [defaultValue.trim(), ...direct];
    }
    return direct;
  }

  const upper = key.toUpperCase();
  if (upper.includes('天數') || upper.includes('VALID_DAYS')) {
    const list = ['7 天', '14 天', '30 天', '45 天', '60 天', '90 天', '180 天', '一年內有效'];
    if (defaultValue && !list.includes(defaultValue.trim()) && defaultValue.trim() !== '') {
      return [defaultValue.trim(), ...list];
    }
    return list;
  }
  if (upper.includes('單位') || upper.includes('UNIT')) {
    const list = ['專案', '式', '套件', '期', '場次', '個月', '小時', '人天', '個', '組', '台', '份'];
    if (defaultValue && !list.includes(defaultValue.trim()) && defaultValue.trim() !== '') {
      return [defaultValue.trim(), ...list];
    }
    return list;
  }
  if (upper.includes('銀行') && (upper.includes('名稱') || upper.includes('BANK_NAME'))) {
    return DROPDOWN_OPTIONS_MAP['受款銀行名稱'];
  }
  return null;
}

/**
 * Parse an ExcelJS workbook for all placeholders.
 * Properly skips slave merged cells to avoid duplicate variable listings.
 */
export function scanWorkbookPlaceholders(workbook: ExcelJS.Workbook): {
  variables: PlaceholderVariable[];
  formulaCount: number;
  imageCount: number;
  totalCellsScanned: number;
  hasHeaderFooter: boolean;
} {
  const varMap = new Map<
    string,
    {
      defaultValue: string;
      occurrences: CellOccurrence[];
      seenAddresses: Set<string>;
    }
  >();

  let formulaCount = 0;
  let totalCellsScanned = 0;
  let imageCount = 0;
  let hasHeaderFooter = false;

  workbook.worksheets.forEach((ws) => {
    try {
      const images = ws.getImages();
      imageCount += images ? images.length : 0;
    } catch {
      // Ignored if unsupported
    }

    if (
      ws.headerFooter &&
      (ws.headerFooter.oddHeader ||
        ws.headerFooter.oddFooter ||
        ws.headerFooter.firstHeader)
    ) {
      hasHeaderFooter = true;
    }

    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        totalCellsScanned++;

        // CRITICAL: Skip slave merged cells to avoid duplicate placeholder scans
        if (cell.isMerged && cell.master && cell.master.address !== cell.address) {
          return;
        }

        // Check if formula cell
        if (
          cell.type === ExcelJS.ValueType.Formula ||
          (typeof cell.value === 'object' &&
            cell.value !== null &&
            'formula' in cell.value)
        ) {
          formulaCount++;
          return;
        }

        // Process cell text
        let textContent = '';
        if (typeof cell.value === 'string') {
          textContent = cell.value;
        } else if (
          cell.value &&
          typeof cell.value === 'object' &&
          'richText' in cell.value
        ) {
          const rich = (cell.value as { richText: Array<{ text: string }> }).richText;
          if (Array.isArray(rich)) {
            textContent = rich.map((r) => r.text || '').join('');
          }
        }

        if (!textContent || !textContent.includes('{{')) return;

        const regex = new RegExp(PLACEHOLDER_REGEX.source, 'g');
        let match: RegExpExecArray | null;
        const displayAddress = cell.address;

        while ((match = regex.exec(textContent)) !== null) {
          const rawKey = match[1]?.trim();
          const rawDefault = match[2]?.trim() || '';

          if (!rawKey) continue;

          const occurrence: CellOccurrence = {
            sheetName: ws.name,
            cellAddress: displayAddress,
            row: rowNumber,
            col: colNumber,
            originalText: textContent,
          };

          const keyLookup = `${ws.name}!${displayAddress}`;

          if (varMap.has(rawKey)) {
            const existing = varMap.get(rawKey)!;
            if (!existing.defaultValue && rawDefault) {
              existing.defaultValue = rawDefault;
            }
            if (!existing.seenAddresses.has(keyLookup)) {
              existing.seenAddresses.add(keyLookup);
              existing.occurrences.push(occurrence);
            }
          } else {
            varMap.set(rawKey, {
              defaultValue: rawDefault,
              occurrences: [occurrence],
              seenAddresses: new Set([keyLookup]),
            });
          }
        }
      });
    });
  });

  const variables: PlaceholderVariable[] = Array.from(varMap.entries()).map(
    ([key, data]) => {
      return {
        key,
        defaultValue: data.defaultValue,
        currentValue: data.defaultValue,
        isLongText: isLongTextKey(key, data.defaultValue),
        isDate: isDateKey(key, data.defaultValue),
        dropdownOptions: getDropdownOptions(key, data.defaultValue) || undefined,
        occurrences: data.occurrences,
        category: categorizeKey(key),
      };
    }
  );

  return {
    variables,
    formulaCount,
    imageCount,
    totalCellsScanned,
    hasHeaderFooter,
  };
}

/**
 * Load workbook from an ArrayBuffer
 */
export async function loadWorkbook(buffer: ArrayBuffer): Promise<ExcelJS.Workbook> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  return workbook;
}

interface MergeModel {
  top: number;
  left: number;
  bottom: number;
  right: number;
  range: string;
}

/**
 * Safely splice rows in ExcelJS while preserving and shifting all merged cells accurately!
 */
export function safeSpliceRows(ws: ExcelJS.Worksheet, startRow: number, deleteCount: number) {
  if (deleteCount <= 0) return;

  // 1. Capture all existing merged cell ranges
  const merges: MergeModel[] = [];
  const rawMerges = (ws as any)._merges || {};
  Object.values(rawMerges).forEach((m: any) => {
    if (m && typeof m.top === 'number') {
      merges.push({
        top: m.top,
        left: m.left,
        bottom: m.bottom,
        right: m.right,
        range: `${String.fromCharCode(64 + m.left)}${m.top}:${String.fromCharCode(64 + m.right)}${m.bottom}`,
      });
    }
  });

  // 2. Unmerge all to prevent ExcelJS internal merge corruption
  merges.forEach((m) => {
    try {
      ws.unMergeCells(m.range);
    } catch {
      // Ignored
    }
  });

  // 3. Perform splice
  ws.spliceRows(startRow, deleteCount);

  // 4. Re-apply shifted merges
  merges.forEach((m) => {
    if (m.top >= startRow + deleteCount) {
      // Shift up
      const newTop = m.top - deleteCount;
      const newBottom = m.bottom - deleteCount;
      try {
        ws.mergeCells(newTop, m.left, newBottom, m.right);
      } catch (e) {
        console.warn('Failed to re-merge shifted range:', e);
      }
    } else if (m.bottom < startRow) {
      // Above splice point: restore unchanged
      try {
        ws.mergeCells(m.top, m.left, m.bottom, m.right);
      } catch (e) {
        console.warn('Failed to re-merge range:', e);
      }
    }
  });
}

/**
 * Safely insert rows in ExcelJS while preserving and shifting all merged cells accurately!
 */
export function safeInsertRows(
  ws: ExcelJS.Worksheet,
  startRow: number,
  insertCount: number,
  styleSourceRow?: number
) {
  if (insertCount <= 0) return;

  const merges: MergeModel[] = [];
  const rawMerges = (ws as any)._merges || {};
  Object.values(rawMerges).forEach((m: any) => {
    if (m && typeof m.top === 'number') {
      merges.push({
        top: m.top,
        left: m.left,
        bottom: m.bottom,
        right: m.right,
        range: `${String.fromCharCode(64 + m.left)}${m.top}:${String.fromCharCode(64 + m.right)}${m.bottom}`,
      });
    }
  });

  merges.forEach((m) => {
    try {
      ws.unMergeCells(m.range);
    } catch {
      // Ignored
    }
  });

  for (let i = 0; i < insertCount; i++) {
    const targetRowIdx = startRow + i;
    ws.insertRow(targetRowIdx, []);
    const insertedRow = ws.getRow(targetRowIdx);

    if (styleSourceRow) {
      const srcRow = ws.getRow(styleSourceRow);
      insertedRow.height = srcRow.height || 22;
      for (let col = 1; col <= 7; col++) {
        const srcCell = srcRow.getCell(col);
        const targetCell = insertedRow.getCell(col);
        if (srcCell.font) targetCell.font = { ...srcCell.font };
        if (srcCell.alignment) targetCell.alignment = { ...srcCell.alignment };
        if (srcCell.border) targetCell.border = { ...srcCell.border };
        if (srcCell.fill) targetCell.fill = { ...srcCell.fill };
        if (srcCell.numFmt) targetCell.numFmt = srcCell.numFmt;
      }
    }
  }

  merges.forEach((m) => {
    if (m.top >= startRow) {
      const newTop = m.top + insertCount;
      const newBottom = m.bottom + insertCount;
      try {
        ws.mergeCells(newTop, m.left, newBottom, m.right);
      } catch (e) {
        console.warn('Failed to re-merge inserted shifted range:', e);
      }
    } else if (m.bottom < startRow) {
      try {
        ws.mergeCells(m.top, m.left, m.bottom, m.right);
      } catch (e) {
        console.warn('Failed to re-merge range:', e);
      }
    }
  });
}

/**
 * Fill values into a cloned workbook and export to file.
 * Preserves 100% of formatting, logo images, formulas, and merged blocks!
 */
export async function exportFilledWorkbook(
  originalBuffer: ArrayBuffer,
  values: Record<string, string>,
  fileName: string,
  options?: {
    embedLogo?: boolean;
    tableMode?: 'auto-shrink' | 'keep-template-rows';
  }
): Promise<Blob> {
  const tableMode = options?.tableMode || 'auto-shrink';
  const workbook = await loadWorkbook(originalBuffer);

  // If embedLogo is enabled, check if /logo.png exists and embed into sheet
  if (options?.embedLogo) {
    try {
      const logoRes = await fetch('/logo.png');
      if (logoRes.ok) {
        const logoBuf = await logoRes.arrayBuffer();
        if (logoBuf.byteLength > 100) {
          const ws = workbook.getWorksheet('商務報價單') || workbook.worksheets[0];
          if (ws) {
            const imgId = workbook.addImage({
              buffer: logoBuf,
              extension: 'png',
            });
            ws.addImage(imgId, {
              tl: { col: 0.1, row: 1.1 },
              ext: { width: 115, height: 48 },
              editAs: 'oneCell',
            });
            // Clear placeholder text in A2:B4 if it had COMPANY_LOGO or 公司商標名稱
            for (let r = 2; r <= 4; r++) {
              for (let c = 1; c <= 2; c++) {
                const cell = ws.getCell(r, c);
                if (
                  typeof cell.value === 'string' &&
                  (cell.value.includes('COMPANY_LOGO') || cell.value.includes('公司商標名稱'))
                ) {
                  cell.value = '';
                }
              }
            }
          }
        }
      }
    } catch (e) {
      console.warn('Failed to embed logo into export:', e);
    }
  }

  workbook.worksheets.forEach((ws) => {
    // 1. Locate item table boundaries
    let headerRowIdx = -1;
    let subtotalRowIdx = -1;
    ws.eachRow((row, r) => {
      const a = row.getCell(1).value;
      if (typeof a === 'string' && (a.includes('項次') || a.includes('序號'))) {
        headerRowIdx = r;
      }
      if (
        typeof a === 'string' &&
        (a.includes('銷售額合計') || a.includes('請款金額合計') || a.includes('未稅 Subtotal'))
      ) {
        subtotalRowIdx = r;
      }
    });

    // 2. Extract active items from values
    const itemMap = new Map<
      number,
      { name: string; desc: string; qty: number; unit: string; price: number }
    >();

    Object.entries(values).forEach(([key, val]) => {
      const m1 = key.match(/^(?:ITEM_|品項)(\d+)_(NAME|名稱|DESC|說明|QTY|數量|UNIT|單位|PRICE|單價)$/i);
      if (m1) {
        const idx = parseInt(m1[1], 10);
        if (!itemMap.has(idx)) {
          itemMap.set(idx, { name: '', desc: '', qty: 1, unit: '式', price: 0 });
        }
        const item = itemMap.get(idx)!;
        const type = m1[2].toUpperCase();
        if (type === 'NAME' || type === '名稱') item.name = val;
        else if (type === 'DESC' || type === '說明') item.desc = val;
        else if (type === 'QTY' || type === '數量') item.qty = Number(val) || 0;
        else if (type === 'UNIT' || type === '單位') item.unit = val;
        else if (type === 'PRICE' || type === '單價') item.price = Number(val) || 0;
      }
      const m2 = key.match(
        /^(?:STAGE_|請款階段)(\d+)_(NAME|項目名稱|DESC|工作內容說明|QTY|數量|UNIT|單位|PRICE|單價)$/i
      );
      if (m2) {
        const idx = parseInt(m2[1], 10);
        if (!itemMap.has(idx)) {
          itemMap.set(idx, { name: '', desc: '', qty: 1, unit: '期', price: 0 });
        }
        const item = itemMap.get(idx)!;
        const type = m2[2].toUpperCase();
        if (type === 'NAME' || type === '項目名稱') item.name = val;
        else if (type === 'DESC' || type === '工作內容說明') item.desc = val;
        else if (type === 'QTY' || type === '數量') item.qty = Number(val) || 0;
        else if (type === 'UNIT' || type === '單位') item.unit = val;
        else if (type === 'PRICE' || type === '單價') item.price = Number(val) || 0;
      }
    });

    const activeItemKeys = Array.from(itemMap.keys()).sort((a, b) => a - b);
    const nonBlankItems = activeItemKeys
      .map((k) => itemMap.get(k)!)
      .filter((it) => it.name && it.name.trim() !== '');

    const activeItems = nonBlankItems.length > 0
      ? nonBlankItems
      : activeItemKeys.map((k) => itemMap.get(k)!);

    // 3. Process item table
    if (headerRowIdx > 0 && subtotalRowIdx > headerRowIdx && activeItems.length > 0) {
      const templateItemCount = subtotalRowIdx - headerRowIdx - 1;
      const targetCount = activeItems.length;

      if (tableMode === 'auto-shrink') {
        if (targetCount < templateItemCount) {
          // Safely delete excess rows with merge preservation
          const deleteCount = templateItemCount - targetCount;
          safeSpliceRows(ws, headerRowIdx + targetCount + 1, deleteCount);
        } else if (targetCount > templateItemCount) {
          // Safely insert extra rows with merge preservation
          const insertCount = targetCount - templateItemCount;
          safeInsertRows(ws, subtotalRowIdx, insertCount, subtotalRowIdx - 1);
        }

        // Populate all active item rows
        for (let i = 0; i < targetCount; i++) {
          const rowIdx = headerRowIdx + 1 + i;
          const it = activeItems[i];
          ws.getCell(`A${rowIdx}`).value = i + 1;
          ws.getCell(`B${rowIdx}`).value = it.name;
          ws.getCell(`C${rowIdx}`).value = it.desc;
          ws.getCell(`D${rowIdx}`).value = it.qty;
          ws.getCell(`E${rowIdx}`).value = it.unit;
          ws.getCell(`F${rowIdx}`).value = it.price;
          ws.getCell(`F${rowIdx}`).numFmt = '"NT$ "#,##0';
          ws.getCell(`G${rowIdx}`).value = { formula: `D${rowIdx}*F${rowIdx}` };
          ws.getCell(`G${rowIdx}`).numFmt = '"NT$ "#,##0';
        }

        // Update Subtotal, VAT, and Grand Total formulas
        const firstItemRow = headerRowIdx + 1;
        const lastItemRow = headerRowIdx + targetCount;
        const newSubtotalRow = headerRowIdx + targetCount + 1;

        const subtotalCell = ws.getCell(newSubtotalRow, 7);
        subtotalCell.value = { formula: `SUM(G${firstItemRow}:G${lastItemRow})` };
        subtotalCell.numFmt = '"NT$ "#,##0';

        const vatRow = newSubtotalRow + 1;
        const vatCell = ws.getCell(vatRow, 7);
        vatCell.value = { formula: `ROUND(G${newSubtotalRow}*0.05, 0)` };
        vatCell.numFmt = '"NT$ "#,##0';

        const grandTotalRow = newSubtotalRow + 2;
        const grandTotalCell = ws.getCell(grandTotalRow, 7);
        grandTotalCell.value = { formula: `G${newSubtotalRow}+G${vatRow}` };
        grandTotalCell.numFmt = '"NT$ "#,##0';
      } else {
        // Keep fixed template row layout: fill active items, clear unused placeholder item rows
        for (let i = 0; i < templateItemCount; i++) {
          const rowIdx = headerRowIdx + 1 + i;
          if (i < targetCount && activeItems[i].name) {
            const it = activeItems[i];
            ws.getCell(`A${rowIdx}`).value = i + 1;
            ws.getCell(`B${rowIdx}`).value = it.name;
            ws.getCell(`C${rowIdx}`).value = it.desc;
            ws.getCell(`D${rowIdx}`).value = it.qty;
            ws.getCell(`E${rowIdx}`).value = it.unit;
            ws.getCell(`F${rowIdx}`).value = it.price;
            ws.getCell(`F${rowIdx}`).numFmt = '"NT$ "#,##0';
            ws.getCell(`G${rowIdx}`).value = { formula: `D${rowIdx}*F${rowIdx}` };
            ws.getCell(`G${rowIdx}`).numFmt = '"NT$ "#,##0';
          } else {
            // Unused row: clear cells
            ws.getCell(`A${rowIdx}`).value = null;
            ws.getCell(`B${rowIdx}`).value = null;
            ws.getCell(`C${rowIdx}`).value = null;
            ws.getCell(`D${rowIdx}`).value = null;
            ws.getCell(`E${rowIdx}`).value = null;
            ws.getCell(`F${rowIdx}`).value = null;
            ws.getCell(`G${rowIdx}`).value = { formula: `IF(D${rowIdx}*F${rowIdx}>0, D${rowIdx}*F${rowIdx}, "")` };
          }
        }
      }
    }

    // 4. Fill in all remaining placeholder variables across Header, Footer, and Other cells
    const activeItemCount = tableMode === 'auto-shrink' ? activeItems.length : (subtotalRowIdx - headerRowIdx - 1);
    ws.eachRow({ includeEmpty: false }, (row, r) => {
      // Skip the line items table rows we already handled
      if (
        headerRowIdx > 0 &&
        r > headerRowIdx &&
        r <= headerRowIdx + Math.max(activeItemCount, 1) + 2
      ) {
        return;
      }

      row.eachCell({ includeEmpty: false }, (cell) => {
        // CRITICAL: Skip slave merged cells so we only modify the master cell
        if (cell.isMerged && cell.master && cell.master.address !== cell.address) {
          return;
        }

        // If it's a formula, keep it untouched
        if (
          cell.type === ExcelJS.ValueType.Formula ||
          (typeof cell.value === 'object' && cell.value !== null && 'formula' in cell.value)
        ) {
          return;
        }

        // Handle string cell
        if (typeof cell.value === 'string') {
          const originalStr = cell.value;
          if (!originalStr.includes('{{')) return;

          const exactMatch = originalStr
            .trim()
            .match(/^\{\{\s*([^\|\}]+?)(?:\s*\|\s*([^\{\}]+?))?\s*\}\}$/);
          if (exactMatch) {
            const key = exactMatch[1].trim();
            const replacement = values[key] !== undefined ? values[key] : '';

            // If pure number, parse to numeric for formula execution
            const trimmedRep = replacement.trim();
            const numericVal = Number(trimmedRep);
            if (trimmedRep !== '' && !isNaN(numericVal) && !isNaN(parseFloat(trimmedRep))) {
              cell.value = numericVal;
            } else {
              cell.value = replacement;
            }
            return;
          }

          // Composite string with embedded placeholders
          const replaced = originalStr.replace(
            new RegExp(PLACEHOLDER_REGEX.source, 'g'),
            (_, key) => {
              const trimmedKey = key.trim();
              return values[trimmedKey] !== undefined ? values[trimmedKey] : '';
            }
          );

          cell.value = replaced;
        } else if (cell.value && typeof cell.value === 'object' && 'richText' in cell.value) {
          const rich = (cell.value as { richText: Array<{ text: string; font?: any }> }).richText;
          if (Array.isArray(rich)) {
            rich.forEach((fragment) => {
              if (fragment.text && fragment.text.includes('{{')) {
                fragment.text = fragment.text.replace(
                  new RegExp(PLACEHOLDER_REGEX.source, 'g'),
                  (_, key) => {
                    const trimmedKey = key.trim();
                    return values[trimmedKey] !== undefined ? values[trimmedKey] : '';
                  }
                );
              }
            });
            cell.value = { richText: rich };
          }
        }
      });
    });
  });

  const outBuffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([outBuffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });

  const baseName = fileName.replace(/\.xlsx$/i, '');
  const dateStr = new Date().toISOString().slice(0, 10);
  const downloadName = `填寫完成_${baseName}_${dateStr}.xlsx`;

  FileSaver.saveAs(blob, downloadName);
  return blob;
}

/**
 * Generate in-memory fallback template in case network fetch fails
 */
export async function generateFallbackTemplate(type: 'quote' | 'invoice'): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = '商務表單系統';

  const COLORS = {
    TABLE_HEADER_BG: 'FF334155',
    TABLE_HEADER_TEXT: 'FFFFFFFF',
    SECTION_HEADER_BG: 'FFE2E8F0',
    SECTION_HEADER_TEXT: 'FF0F172A',
    LABEL_BG: 'FFF1F5F9',
    LABEL_TEXT: 'FF334155',
    SUBTOTAL_BG: 'FFF8FAFC',
    TOTAL_BG: 'FFE2E8F0',
    TOTAL_TEXT: 'FF0F172A',
    BORDER_MEDIUM: 'FFCBD5E1',
    BORDER_DARK: 'FF94A3B8',
    TEXT_MAIN: 'FF0F172A',
    TEXT_MUTED: 'FF475569',
  };

  const borderThin = {
    top: { style: 'thin' as const, color: { argb: COLORS.BORDER_MEDIUM } },
    left: { style: 'thin' as const, color: { argb: COLORS.BORDER_MEDIUM } },
    bottom: { style: 'thin' as const, color: { argb: COLORS.BORDER_MEDIUM } },
    right: { style: 'thin' as const, color: { argb: COLORS.BORDER_MEDIUM } },
  };

  if (type === 'quote') {
    const ws = wb.addWorksheet('商務報價單', { views: [{ showGridLines: true }] });
    ws.columns = [
      { key: 'A', width: 8 },
      { key: 'B', width: 22 },
      { key: 'C', width: 34 },
      { key: 'D', width: 12 },
      { key: 'E', width: 12 },
      { key: 'F', width: 18 },
      { key: 'G', width: 22 },
    ];

    ws.mergeCells('A1:G1');
    ws.getCell('A1').value = 'QUOTATION 商務報價單';
    ws.getCell('A1').font = { name: 'Microsoft JhengHei', size: 16, bold: true, color: { argb: COLORS.TEXT_MAIN } };
    ws.getCell('A1').alignment = { vertical: 'middle', horizontal: 'center' };
    ws.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
    ws.getRow(1).height = 36;

    ws.mergeCells('A2:B4');
    ws.getCell('A2').value = '{{公司商標名稱 | ⭐ 創想數位科技}}';
    ws.getCell('A2').font = { name: 'Microsoft JhengHei', size: 13, bold: true, color: { argb: COLORS.TEXT_MAIN } };
    ws.getCell('A2').alignment = { vertical: 'middle', horizontal: 'center' };
    ws.getCell('A2').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.LABEL_BG } };
    ws.getCell('A2').border = borderThin;

    ws.getCell('C2').value = '報價單號：';
    ws.getCell('D2').value = '{{報價單號 | QT-20260901}}';
    ws.mergeCells('D2:E2');
    ws.getCell('F2').value = '報價日期：';
    ws.getCell('G2').value = '{{報價日期 | 2026-09-30}}';

    ws.getCell('C3').value = '專案名稱：';
    ws.getCell('D3').value = '{{專案名稱 | 企業雲端 ERP 升級建置案}}';
    ws.mergeCells('D3:E3');
    ws.getCell('F3').value = '有效天數：';
    ws.getCell('G3').value = '{{有效天數 | 30 天}}';

    ws.getCell('C4').value = '聯絡窗口：';
    ws.getCell('D4').value = '{{客戶聯絡窗口 | 陳志明 專案總監}}';
    ws.mergeCells('D4:E4');
    ws.getCell('F4').value = '付款幣別：';
    ws.getCell('G4').value = '新台幣 (NTD)';

    for (let r = 2; r <= 4; r++) {
      ws.getRow(r).height = 22;
      ['C', 'F'].forEach((c) => {
        const cell = ws.getCell(`${c}${r}`);
        cell.font = { name: 'Microsoft JhengHei', size: 9.5, bold: true, color: { argb: COLORS.LABEL_TEXT } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.LABEL_BG } };
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
        cell.border = borderThin;
      });
      ['D', 'G'].forEach((c) => {
        const cell = ws.getCell(`${c}${r}`);
        cell.font = { name: 'Microsoft JhengHei', size: 9.5, color: { argb: COLORS.TEXT_MAIN } };
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
        cell.border = borderThin;
      });
    }

    // Headers
    ['項次', '項目名稱', '規格說明與工作內容', '數量', '單位', '單價 (NTD)', '小計金額 (NTD)'].forEach((h, i) => {
      const colLetter = String.fromCharCode(65 + i);
      const c = ws.getCell(`${colLetter}12`);
      c.value = h;
      c.font = { name: 'Microsoft JhengHei', size: 10, bold: true, color: { argb: COLORS.TABLE_HEADER_TEXT } };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.TABLE_HEADER_BG } };
      c.alignment = { vertical: 'middle', horizontal: i === 0 || i === 3 || i === 4 ? 'center' : i >= 5 ? 'right' : 'left' };
      c.border = borderThin;
    });
    ws.getRow(12).height = 26;

    const items = [
      { r: 13, n: '{{品項1_名稱 | 企業官網響應式切版}}', s: '{{品項1_說明 | React + TypeScript 現代化 RWD 視覺架構}}', q: '{{品項1_數量 | 1}}', u: '{{品項1_單位 | 專案}}', p: '{{品項1_單價 | 65000}}' },
      { r: 14, n: '{{品項2_名稱 | 純前端 Excel 匯出模組}}', s: '{{品項2_說明 | ExcelJS 樣式渲染、多格式相容與離線下載}}', q: '{{品項2_數量 | 1}}', u: '{{品項2_單位 | 套件}}', p: '{{品項2_單價 | 38000}}' },
    ];

    items.forEach((it, idx) => {
      ws.getRow(it.r).height = 24;
      ws.getCell(`A${it.r}`).value = idx + 1;
      ws.getCell(`B${it.r}`).value = it.n;
      ws.getCell(`C${it.r}`).value = it.s;
      ws.getCell(`D${it.r}`).value = it.q;
      ws.getCell(`E${it.r}`).value = it.u;
      ws.getCell(`F${it.r}`).value = it.p;
      ws.getCell(`G${it.r}`).value = { formula: `D${it.r}*F${it.r}` };
      for (let c = 1; c <= 7; c++) {
        const cell = ws.getRow(it.r).getCell(c);
        cell.font = { name: 'Microsoft JhengHei', size: 9.5, color: { argb: COLORS.TEXT_MAIN } };
        cell.border = borderThin;
        cell.alignment = { vertical: 'middle', horizontal: c === 1 || c === 4 || c === 5 ? 'center' : c >= 6 ? 'right' : 'left' };
      }
    });

    const buf = await wb.xlsx.writeBuffer();
    return buf as ArrayBuffer;
  } else {
    const ws = wb.addWorksheet('商業請款單', { views: [{ showGridLines: true }] });
    ws.columns = [
      { key: 'A', width: 8 },
      { key: 'B', width: 22 },
      { key: 'C', width: 34 },
      { key: 'D', width: 12 },
      { key: 'E', width: 12 },
      { key: 'F', width: 18 },
      { key: 'G', width: 22 },
    ];

    ws.mergeCells('A1:G1');
    ws.getCell('A1').value = 'COMMERCIAL INVOICE 商業請款單';
    ws.getCell('A1').font = { name: 'Microsoft JhengHei', size: 16, bold: true, color: { argb: COLORS.TEXT_MAIN } };
    ws.getCell('A1').alignment = { vertical: 'middle', horizontal: 'center' };
    ws.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
    ws.getRow(1).height = 36;

    const buf = await wb.xlsx.writeBuffer();
    return buf as ArrayBuffer;
  }
}

/**
 * Extract structured line items from variables and values for quotation/invoice tables
 */
export function extractQuotationItems(
  variables: PlaceholderVariable[],
  values: Record<string, string>
): QuotationItem[] {
  const itemMap = new Map<
    number,
    {
      prefix: string;
      nameKey: string;
      descKey: string;
      qtyKey: string;
      unitKey: string;
      priceKey: string;
    }
  >();

  variables.forEach((v) => {
    const match1 = v.key.match(/^品項(\d+)_(.+)$/);
    if (match1) {
      const idx = parseInt(match1[1], 10);
      if (!itemMap.has(idx)) {
        itemMap.set(idx, {
          prefix: `品項${idx}`,
          nameKey: `品項${idx}_名稱`,
          descKey: `品項${idx}_說明`,
          qtyKey: `品項${idx}_數量`,
          unitKey: `品項${idx}_單位`,
          priceKey: `品項${idx}_單價`,
        });
      }
      return;
    }

    const match2 = v.key.match(/^請款階段(\d+)_(.+)$/);
    if (match2) {
      const idx = parseInt(match2[1], 10);
      if (!itemMap.has(idx)) {
        itemMap.set(idx, {
          prefix: `請款階段${idx}`,
          nameKey: `請款階段${idx}_項目名稱`,
          descKey: `請款階段${idx}_工作內容說明`,
          qtyKey: `請款階段${idx}_數量`,
          unitKey: `請款階段${idx}_單位`,
          priceKey: `請款階段${idx}_單價`,
        });
      }
      return;
    }

    const match3 = v.key.match(/^ITEM_(\d+)_(.+)$/i);
    if (match3) {
      const idx = parseInt(match3[1], 10);
      if (!itemMap.has(idx)) {
        itemMap.set(idx, {
          prefix: `ITEM_${idx}`,
          nameKey: `ITEM_${idx}_NAME`,
          descKey: `ITEM_${idx}_DESC`,
          qtyKey: `ITEM_${idx}_QTY`,
          unitKey: `ITEM_${idx}_UNIT`,
          priceKey: `ITEM_${idx}_PRICE`,
        });
      }
      return;
    }
  });

  const sortedIndices = Array.from(itemMap.keys()).sort((a, b) => a - b);

  return sortedIndices.map((idx) => {
    const meta = itemMap.get(idx)!;
    const nameVar = variables.find((v) => v.key === meta.nameKey);
    const descVar = variables.find((v) => v.key === meta.descKey);
    const qtyVar = variables.find((v) => v.key === meta.qtyKey);
    const unitVar = variables.find((v) => v.key === meta.unitKey);
    const priceVar = variables.find((v) => v.key === meta.priceKey);

    return {
      id: `item-${idx}`,
      rowNumber: idx,
      nameKey: meta.nameKey,
      descKey: meta.descKey,
      qtyKey: meta.qtyKey,
      unitKey: meta.unitKey,
      priceKey: meta.priceKey,
      name: values[meta.nameKey] !== undefined ? values[meta.nameKey] : (nameVar?.defaultValue || ''),
      desc: values[meta.descKey] !== undefined ? values[meta.descKey] : (descVar?.defaultValue || ''),
      qty: values[meta.qtyKey] !== undefined ? values[meta.qtyKey] : (qtyVar?.defaultValue || '1'),
      unit: values[meta.unitKey] !== undefined ? values[meta.unitKey] : (unitVar?.defaultValue || '式'),
      price: values[meta.priceKey] !== undefined ? values[meta.priceKey] : (priceVar?.defaultValue || '0'),
    };
  });
}

/**
 * Dynamically create the next quotation item and register placeholder variables
 */
export function createNextItem(
  existingItems: QuotationItem[],
  variables: PlaceholderVariable[]
): {
  newItem: QuotationItem;
  newVariables: PlaceholderVariable[];
  newValues: Record<string, string>;
} {
  const maxIdx = existingItems.length > 0 ? Math.max(...existingItems.map((i) => i.rowNumber)) : 0;
  const nextIdx = maxIdx + 1;

  const isInvoice = existingItems.some((i) => i.nameKey.includes('請款階段') || i.nameKey.includes('STAGE_'));
  const isEnPrefix = existingItems.some((i) => i.nameKey.startsWith('ITEM_'));

  const nameKey = isInvoice
    ? `請款階段${nextIdx}_項目名稱`
    : isEnPrefix
    ? `ITEM_${nextIdx}_NAME`
    : `品項${nextIdx}_名稱`;

  const descKey = isInvoice
    ? `請款階段${nextIdx}_工作內容說明`
    : isEnPrefix
    ? `ITEM_${nextIdx}_DESC`
    : `品項${nextIdx}_說明`;

  const qtyKey = isInvoice
    ? `請款階段${nextIdx}_數量`
    : isEnPrefix
    ? `ITEM_${nextIdx}_QTY`
    : `品項${nextIdx}_數量`;

  const unitKey = isInvoice
    ? `請款階段${nextIdx}_單位`
    : isEnPrefix
    ? `ITEM_${nextIdx}_UNIT`
    : `品項${nextIdx}_單位`;

  const priceKey = isInvoice
    ? `請款階段${nextIdx}_單價`
    : isEnPrefix
    ? `ITEM_${nextIdx}_PRICE`
    : `品項${nextIdx}_單價`;

  const newItem: QuotationItem = {
    id: `item-${nextIdx}-${Date.now()}`,
    rowNumber: nextIdx,
    nameKey,
    descKey,
    qtyKey,
    unitKey,
    priceKey,
    name: `新增品項 ${nextIdx}`,
    desc: '規格與功能說明',
    qty: '1',
    unit: '式',
    price: '0',
  };

  const newValues: Record<string, string> = {
    [nameKey]: newItem.name,
    [descKey]: newItem.desc,
    [qtyKey]: newItem.qty,
    [unitKey]: newItem.unit,
    [priceKey]: newItem.price,
  };

  const newVariables: PlaceholderVariable[] = [
    {
      key: nameKey,
      defaultValue: newItem.name,
      currentValue: newItem.name,
      isLongText: false,
      occurrences: [],
      category: '明細項目',
    },
    {
      key: descKey,
      defaultValue: newItem.desc,
      currentValue: newItem.desc,
      isLongText: true,
      occurrences: [],
      category: '明細項目',
    },
    {
      key: qtyKey,
      defaultValue: newItem.qty,
      currentValue: newItem.qty,
      isLongText: false,
      occurrences: [],
      category: '明細項目',
    },
    {
      key: unitKey,
      defaultValue: newItem.unit,
      currentValue: newItem.unit,
      isLongText: false,
      dropdownOptions: ['專案', '式', '套件', '期', '場次', '個月', '小時', '人天', '個', '組', '台', '份'],
      occurrences: [],
      category: '明細項目',
    },
    {
      key: priceKey,
      defaultValue: newItem.price,
      currentValue: newItem.price,
      isLongText: false,
      occurrences: [],
      category: '明細項目',
    },
  ];

  return { newItem, newVariables, newValues };
}
