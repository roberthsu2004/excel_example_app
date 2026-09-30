const ExcelJS = require('exceljs');

const COLORS = {
  TABLE_HEADER_BG: 'FF1E293B',   // Dark navy / slate 800
  TABLE_HEADER_TEXT: 'FFFFFFFF', // Pure White
  SECTION_HEADER_BG: 'FFF8FAFC', // Slate 50
  SECTION_HEADER_TEXT: 'FF0F172A', // Slate 900
  TOTAL_BG: 'FFE8ECF2',          // Subtle light grey
  TOTAL_TEXT: 'FF0F172A',        // Slate 900
  BORDER_GREY: 'FFCBD5E1',       // Slate 300
  BORDER_DARK: 'FF475569',       // Slate 600
  TEXT_MAIN: 'FF0F172A',         // Slate 900
  TEXT_MUTED: 'FF475569',        // Slate 600
  TEXT_LIGHT: 'FF64748B',        // Slate 500
};

const BORDER_STYLE = {
  top: { style: 'thin', color: { argb: COLORS.BORDER_GREY } },
  left: { style: 'thin', color: { argb: COLORS.BORDER_GREY } },
  bottom: { style: 'thin', color: { argb: COLORS.BORDER_GREY } },
  right: { style: 'thin', color: { argb: COLORS.BORDER_GREY } },
};

const BORDER_HEADER = {
  top: { style: 'thin', color: { argb: COLORS.BORDER_DARK } },
  left: { style: 'thin', color: { argb: COLORS.BORDER_DARK } },
  bottom: { style: 'thin', color: { argb: COLORS.BORDER_DARK } },
  right: { style: 'thin', color: { argb: COLORS.BORDER_DARK } },
};

function applyCellRangeStyle(ws, topRow, leftCol, bottomRow, rightCol, styleFn) {
  for (let r = topRow; r <= bottomRow; r++) {
    for (let c = leftCol; c <= rightCol; c++) {
      const cell = ws.getCell(r, c);
      styleFn(cell, r, c);
    }
  }
}

async function buildQuotationTemplate() {
  const wb = new ExcelJS.Workbook();
  wb.creator = '商務表單系統';
  const ws = wb.addWorksheet('商務報價單', {
    views: [{ showGridLines: true }],
  });

  // Column widths matching 100% with screenshot
  ws.columns = [
    { key: 'A', width: 6.5 }, // 項次
    { key: 'B', width: 23 },  // 項目名稱
    { key: 'C', width: 38 },  // 規格說明與功能簡述
    { key: 'D', width: 8.5 }, // 數量
    { key: 'E', width: 11 },  // 單位
    { key: 'F', width: 18 },  // 單價 (NTD)
    { key: 'G', width: 26 },  // 小計金額 (NTD) - set to 26 so NT$ 190,050 never shows ###
  ];

  // Row 1: Blank separator
  ws.getRow(1).height = 14;

  // Row 2-4: Left Box (A2:B4) - Company Logo
  ws.mergeCells('A2:B4');
  const logoCell = ws.getCell('A2');
  logoCell.value = '❖\n{{COMPANY_LOGO | ACME\nCORP 智慧科技}}';
  applyCellRangeStyle(ws, 2, 1, 4, 2, (cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
    cell.font = { name: 'Microsoft JhengHei', size: 10.5, bold: true, color: { argb: COLORS.TEXT_MAIN } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = BORDER_STYLE;
  });

  // Row 2: Title (C2:G2)
  ws.mergeCells('C2:G2');
  const titleCell = ws.getCell('C2');
  titleCell.value = '商 務 報 價 單  /  QUOTATION';
  applyCellRangeStyle(ws, 2, 3, 2, 7, (cell) => {
    cell.font = { name: 'Microsoft JhengHei', size: 16, bold: true, color: { argb: COLORS.TEXT_MAIN } };
    cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  });
  ws.getRow(2).height = 28;

  // Row 3: Metadata (C3:G3) - Quotation No & Date
  ws.mergeCells('C3:G3');
  const metaCell1 = ws.getCell('C3');
  metaCell1.value = '報價單號：{{QUOTATION_NO | QT-20260926-001}}  ｜  報價日期：{{QUOTATION_DATE | 2026-09-26}}';
  applyCellRangeStyle(ws, 3, 3, 3, 7, (cell) => {
    cell.font = { name: 'Microsoft JhengHei', size: 9.5, color: { argb: COLORS.TEXT_MUTED } };
    cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    cell.border = BORDER_STYLE;
  });
  ws.getRow(3).height = 20;

  // Row 4: Metadata (C4:G4) - Project Name
  ws.mergeCells('C4:G4');
  const metaCell2 = ws.getCell('C4');
  metaCell2.value = '專案名稱：{{PROJECT_NAME | 企業數位轉型前端自動化與報價管理系統建置}}';
  applyCellRangeStyle(ws, 4, 3, 4, 7, (cell) => {
    cell.font = { name: 'Microsoft JhengHei', size: 9.5, color: { argb: COLORS.TEXT_MUTED } };
    cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    cell.border = BORDER_STYLE;
  });
  ws.getRow(4).height = 20;

  // Row 5: Blank separator
  ws.getRow(5).height = 12;

  // Row 6: Section Headers (A6:D6 & E6:G6)
  ws.mergeCells('A6:D6');
  ws.getCell('A6').value = '【 供應商資訊 (我方) 】';
  applyCellRangeStyle(ws, 6, 1, 6, 4, (cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.SECTION_HEADER_BG } };
    cell.font = { name: 'Microsoft JhengHei', size: 9.5, bold: true, color: { argb: COLORS.SECTION_HEADER_TEXT } };
    cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    cell.border = BORDER_STYLE;
  });

  ws.mergeCells('E6:G6');
  ws.getCell('E6').value = '【 客戶買受人資訊 】';
  applyCellRangeStyle(ws, 6, 5, 6, 7, (cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.SECTION_HEADER_BG } };
    cell.font = { name: 'Microsoft JhengHei', size: 9.5, bold: true, color: { argb: COLORS.SECTION_HEADER_TEXT } };
    cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    cell.border = BORDER_STYLE;
  });
  ws.getRow(6).height = 22;

  // Rows 7-10: Provider & Client Information Fields
  const partyRows = [
    {
      r: 7,
      left: '公司名稱：{{PROVIDER_NAME | 星創智慧科技有限公司}}',
      right: '客戶名稱：{{CLIENT_NAME | 宏達數位傳媒行銷有限公司}}',
    },
    {
      r: 8,
      left: '統一編號：{{PROVIDER_TAX_ID | 88889999}}',
      right: '統一編號：{{CLIENT_TAX_ID | 12345678}}',
    },
    {
      r: 9,
      left: '聯絡電話：{{PROVIDER_PHONE | (02) 2345-6789}}',
      right: '聯絡窗口：{{CLIENT_CONTACT | 陳專案經理 (0912-345-678)}}',
    },
    {
      r: 10,
      left: '公司地址：{{PROVIDER_ADDRESS | 台北市信義區信義路五段7號}}',
      right: '通訊地址：{{CLIENT_ADDRESS | 台北市中山區南京東路二段100號}}',
    },
  ];

  partyRows.forEach(item => {
    ws.getRow(item.r).height = 20;

    // Left info (A:D)
    ws.mergeCells(`A${item.r}:D${item.r}`);
    ws.getCell(`A${item.r}`).value = item.left;
    applyCellRangeStyle(ws, item.r, 1, item.r, 4, (cell) => {
      cell.font = { name: 'Microsoft JhengHei', size: 9, color: { argb: COLORS.TEXT_MAIN } };
      cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
      cell.border = BORDER_STYLE;
    });

    // Right info (E:G)
    ws.mergeCells(`E${item.r}:G${item.r}`);
    ws.getCell(`E${item.r}`).value = item.right;
    applyCellRangeStyle(ws, item.r, 5, item.r, 7, (cell) => {
      cell.font = { name: 'Microsoft JhengHei', size: 9, color: { argb: COLORS.TEXT_MAIN } };
      cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
      cell.border = BORDER_STYLE;
    });
  });

  // Row 11: Blank separator
  ws.getRow(11).height = 12;

  // Row 12: Table Header
  const headers = [
    { col: 'A', text: '項次', align: 'center' },
    { col: 'B', text: '項目名稱', align: 'center' },
    { col: 'C', text: '規格說明與功能簡述', align: 'center' },
    { col: 'D', text: '數量', align: 'center' },
    { col: 'E', text: '單位', align: 'center' },
    { col: 'F', text: '單價 (NTD)', align: 'center' },
    { col: 'G', text: '小計金額 (NTD)', align: 'center' },
  ];

  headers.forEach(h => {
    const c = ws.getCell(`${h.col}12`);
    c.value = h.text;
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.TABLE_HEADER_BG } };
    c.font = { name: 'Microsoft JhengHei', size: 9.5, bold: true, color: { argb: COLORS.TABLE_HEADER_TEXT } };
    c.alignment = { vertical: 'middle', horizontal: h.align };
    c.border = BORDER_HEADER;
  });
  ws.getRow(12).height = 24;

  // Rows 13-17: Item Rows
  const items = [
    {
      r: 13,
      no: 1,
      name: '{{ITEM_1_NAME | 企業官網響應式切版}}',
      desc: '{{ITEM_1_DESC | React + TypeScript 現代化 RWD 視覺架構}}',
      qty: '{{ITEM_1_QTY | 1}}',
      unit: '{{ITEM_1_UNIT | 專案}}',
      price: '{{ITEM_1_PRICE | 65000}}',
    },
    {
      r: 14,
      no: 2,
      name: '{{ITEM_2_NAME | 純前端 Excel 匯出模組}}',
      desc: '{{ITEM_2_DESC | ExcelJS 樣式渲染、多格式相容與離線下載}}',
      qty: '{{ITEM_2_QTY | 1}}',
      unit: '{{ITEM_2_UNIT | 套件}}',
      price: '{{ITEM_2_PRICE | 38000}}',
    },
    {
      r: 15,
      no: 3,
      name: '{{ITEM_3_NAME | 商務邏輯與表單運算}}',
      desc: '{{ITEM_3_DESC | 營業稅、千分位計算、動態增刪項目與防呆}}',
      qty: '{{ITEM_3_QTY | 2}}',
      unit: '{{ITEM_3_UNIT | 式}}',
      price: '{{ITEM_3_PRICE | 15000}}',
    },
    {
      r: 16,
      no: 4,
      name: '{{ITEM_4_NAME | 系統教育訓練與交接}}',
      desc: '{{ITEM_4_DESC | 提供完整原始碼、操作手冊與 4 小時講師訓練}}',
      qty: '{{ITEM_4_QTY | 1}}',
      unit: '{{ITEM_4_UNIT | 場次}}',
      price: '{{ITEM_4_PRICE | 12000}}',
    },
    {
      r: 17,
      no: 5,
      name: '{{ITEM_5_NAME | 首年維護與技術諮詢}}',
      desc: '{{ITEM_5_DESC | 日常功能保固、相依套件升級檢視與線上諮詢}}',
      qty: '{{ITEM_5_QTY | 12}}',
      unit: '{{ITEM_5_UNIT | 個月}}',
      price: '{{ITEM_5_PRICE | 3000}}',
    },
  ];

  items.forEach(it => {
    ws.getRow(it.r).height = 22;

    // A: Index (Center)
    const cA = ws.getCell(`A${it.r}`);
    cA.value = it.no;
    cA.alignment = { vertical: 'middle', horizontal: 'center' };
    cA.font = { name: 'Microsoft JhengHei', size: 9.5, color: { argb: COLORS.TEXT_MAIN } };
    cA.border = BORDER_STYLE;

    // B: Name (Left)
    const cB = ws.getCell(`B${it.r}`);
    cB.value = it.name;
    cB.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    cB.font = { name: 'Microsoft JhengHei', size: 9.5, color: { argb: COLORS.TEXT_MAIN } };
    cB.border = BORDER_STYLE;

    // C: Desc (Left)
    const cC = ws.getCell(`C${it.r}`);
    cC.value = it.desc;
    cC.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    cC.font = { name: 'Microsoft JhengHei', size: 9, color: { argb: COLORS.TEXT_MUTED } };
    cC.border = BORDER_STYLE;

    // D: Qty (Center)
    const cD = ws.getCell(`D${it.r}`);
    cD.value = it.qty;
    cD.alignment = { vertical: 'middle', horizontal: 'center' };
    cD.font = { name: 'Microsoft JhengHei', size: 9.5, color: { argb: COLORS.TEXT_MAIN } };
    cD.border = BORDER_STYLE;

    // E: Unit (Center)
    const cE = ws.getCell(`E${it.r}`);
    cE.value = it.unit;
    cE.alignment = { vertical: 'middle', horizontal: 'center' };
    cE.font = { name: 'Microsoft JhengHei', size: 9.5, color: { argb: COLORS.TEXT_MAIN } };
    cE.border = BORDER_STYLE;

    // F: Price (Right)
    const cF = ws.getCell(`F${it.r}`);
    cF.value = it.price;
    cF.alignment = { vertical: 'middle', horizontal: 'right' };
    cF.font = { name: 'Microsoft JhengHei', size: 9.5, color: { argb: COLORS.TEXT_MAIN } };
    cF.numFmt = '"NT$ "#,##0';
    cF.border = BORDER_STYLE;

    // G: Subtotal formula (Right)
    const cG = ws.getCell(`G${it.r}`);
    cG.value = { formula: `D${it.r}*F${it.r}` };
    cG.alignment = { vertical: 'middle', horizontal: 'right' };
    cG.font = { name: 'Microsoft JhengHei', size: 9.5, color: { argb: COLORS.TEXT_MAIN } };
    cG.numFmt = '"NT$ "#,##0';
    cG.border = BORDER_STYLE;
  });

  // Row 18: Subtotal
  ws.mergeCells('A18:F18');
  ws.getCell('A18').value = '銷售額合計 (未稅 Subtotal)';
  applyCellRangeStyle(ws, 18, 1, 18, 6, (cell) => {
    cell.font = { name: 'Microsoft JhengHei', size: 9.5, color: { argb: COLORS.TEXT_MAIN } };
    cell.alignment = { vertical: 'middle', horizontal: 'right' };
    cell.border = BORDER_STYLE;
  });
  const cG18 = ws.getCell('G18');
  cG18.value = { formula: 'SUM(G13:G17)' };
  cG18.font = { name: 'Microsoft JhengHei', size: 9.5, color: { argb: COLORS.TEXT_MAIN } };
  cG18.alignment = { vertical: 'middle', horizontal: 'right' };
  cG18.numFmt = '"NT$ "#,##0';
  cG18.border = BORDER_STYLE;
  ws.getRow(18).height = 22;

  // Row 19: VAT 5%
  ws.mergeCells('A19:F19');
  ws.getCell('A19').value = '加值型營業稅 (VAT 5%)';
  applyCellRangeStyle(ws, 19, 1, 19, 6, (cell) => {
    cell.font = { name: 'Microsoft JhengHei', size: 9.5, color: { argb: COLORS.TEXT_MAIN } };
    cell.alignment = { vertical: 'middle', horizontal: 'right' };
    cell.border = BORDER_STYLE;
  });
  const cG19 = ws.getCell('G19');
  cG19.value = { formula: 'ROUND(G18*0.05, 0)' };
  cG19.font = { name: 'Microsoft JhengHei', size: 9.5, color: { argb: COLORS.TEXT_MAIN } };
  cG19.alignment = { vertical: 'middle', horizontal: 'right' };
  cG19.numFmt = '"NT$ "#,##0';
  cG19.border = BORDER_STYLE;
  ws.getRow(19).height = 22;

  // Row 20: Grand Total
  ws.mergeCells('A20:F20');
  ws.getCell('A20').value = '總計金額 (含稅 Grand Total NTD)';
  applyCellRangeStyle(ws, 20, 1, 20, 6, (cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.TOTAL_BG } };
    cell.font = { name: 'Microsoft JhengHei', size: 10, bold: true, color: { argb: COLORS.TOTAL_TEXT } };
    cell.alignment = { vertical: 'middle', horizontal: 'right' };
    cell.border = {
      top: { style: 'thin', color: { argb: COLORS.BORDER_GREY } },
      left: { style: 'thin', color: { argb: COLORS.BORDER_GREY } },
      bottom: { style: 'double', color: { argb: COLORS.TEXT_MAIN } },
      right: { style: 'thin', color: { argb: COLORS.BORDER_GREY } },
    };
  });
  const cG20 = ws.getCell('G20');
  cG20.value = { formula: 'G18+G19' };
  cG20.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.TOTAL_BG } };
  cG20.font = { name: 'Microsoft JhengHei', size: 10.5, bold: true, color: { argb: 'FF1E3A8A' } };
  cG20.alignment = { vertical: 'middle', horizontal: 'right' };
  cG20.numFmt = '"NT$ "#,##0';
  cG20.border = {
    top: { style: 'thin', color: { argb: COLORS.BORDER_GREY } },
    left: { style: 'thin', color: { argb: COLORS.BORDER_GREY } },
    bottom: { style: 'double', color: { argb: COLORS.TEXT_MAIN } },
    right: { style: 'thin', color: { argb: COLORS.BORDER_GREY } },
  };
  ws.getRow(20).height = 24;

  // Row 21: Blank separator
  ws.getRow(21).height = 12;

  // Row 22: Terms Section Header
  ws.mergeCells('A22:G22');
  ws.getCell('A22').value = '【 交易條款與備註說明 】';
  applyCellRangeStyle(ws, 22, 1, 22, 7, (cell) => {
    cell.font = { name: 'Microsoft JhengHei', size: 9.5, bold: true, color: { argb: COLORS.TEXT_MAIN } };
    cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    cell.border = BORDER_STYLE;
  });
  ws.getRow(22).height = 22;

  // Rows 23-25: Terms items
  const terms = [
    { r: 23, text: '1. 付款條件：{{PAYMENT_TERMS | 本專案依階段請款，簽約訂金 30%、期中展示驗收 40%、上線結案驗收 30%。}}' },
    { r: 24, text: '2. 匯款專戶：{{BANK_ACCOUNT | 玉山銀行 (808) 營業部  帳號：1234-567-890123  戶名：星創智慧科技有限公司}}' },
    { r: 25, text: '3. 有效天數：{{VALID_TERMS | 本報價單內容自開立日起 30 日內有效，若逾期未確認需重新評估報價。}}' },
  ];

  terms.forEach(t => {
    ws.getRow(t.r).height = 20;
    ws.mergeCells(`A${t.r}:G${t.r}`);
    ws.getCell(`A${t.r}`).value = t.text;
    applyCellRangeStyle(ws, t.r, 1, t.r, 7, (cell) => {
      cell.font = { name: 'Microsoft JhengHei', size: 9, color: { argb: COLORS.TEXT_MUTED } };
      cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
      cell.border = BORDER_STYLE;
    });
  });

  // Row 26: Blank separator
  ws.getRow(26).height = 12;

  // Row 27: Signature Box Headers
  ws.mergeCells('A27:C27');
  ws.getCell('A27').value = '供應商確認簽章 (發票章 / 大小章)';
  applyCellRangeStyle(ws, 27, 1, 27, 3, (cell) => {
    cell.font = { name: 'Microsoft JhengHei', size: 9, color: { argb: COLORS.TEXT_MUTED } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = BORDER_STYLE;
  });

  // Cell D27
  ws.getCell('D27').border = BORDER_STYLE;

  ws.mergeCells('E27:G27');
  ws.getCell('E27').value = '客戶簽名確認回傳 (請簽名並加蓋公司用印)';
  applyCellRangeStyle(ws, 27, 5, 27, 7, (cell) => {
    cell.font = { name: 'Microsoft JhengHei', size: 9, color: { argb: COLORS.TEXT_MUTED } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = BORDER_STYLE;
  });
  ws.getRow(27).height = 20;

  // Rows 28-30: Signature box body
  ws.mergeCells('A28:C30');
  ws.getCell('A28').value = '\n（請在此處蓋用印章）\n日期：   年   月   日';
  applyCellRangeStyle(ws, 28, 1, 30, 3, (cell) => {
    cell.font = { name: 'Microsoft JhengHei', size: 8.5, color: { argb: COLORS.TEXT_LIGHT } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = BORDER_STYLE;
  });

  // D28, D29, D30 with thin borders
  for (let r = 28; r <= 30; r++) {
    ws.getCell(`D${r}`).border = BORDER_STYLE;
    ws.getRow(r).height = 18;
  }

  ws.mergeCells('E28:G30');
  ws.getCell('E28').value = '\n（請在此處簽署並加蓋公司用印）\n日期：   年   月   日';
  applyCellRangeStyle(ws, 28, 5, 30, 7, (cell) => {
    cell.font = { name: 'Microsoft JhengHei', size: 8.5, color: { argb: COLORS.TEXT_LIGHT } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = BORDER_STYLE;
  });

  await wb.xlsx.writeFile('public/商務報價單範本.xlsx');
  console.log('Saved 100% matched public/商務報價單範本.xlsx');
}

async function buildInvoiceTemplate() {
  const wb = new ExcelJS.Workbook();
  wb.creator = '商務表單系統';
  const ws = wb.addWorksheet('商業請款單', {
    views: [{ showGridLines: true }],
  });

  ws.columns = [
    { key: 'A', width: 6.5 },
    { key: 'B', width: 23 },
    { key: 'C', width: 38 },
    { key: 'D', width: 8.5 },
    { key: 'E', width: 11 },
    { key: 'F', width: 18 },
    { key: 'G', width: 26 },
  ];

  // Row 1: Blank separator
  ws.getRow(1).height = 14;

  // Row 2-4: Left Box (A2:B4) - Company Logo
  ws.mergeCells('A2:B4');
  const logoCell = ws.getCell('A2');
  logoCell.value = '❖\n{{COMPANY_LOGO | 智創雲端\n系統整合科技}}';
  applyCellRangeStyle(ws, 2, 1, 4, 2, (cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
    cell.font = { name: 'Microsoft JhengHei', size: 10.5, bold: true, color: { argb: COLORS.TEXT_MAIN } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = BORDER_STYLE;
  });

  // Row 2: Title (C2:G2)
  ws.mergeCells('C2:G2');
  const titleCell = ws.getCell('C2');
  titleCell.value = '商 業 請 款 單  /  INVOICE';
  applyCellRangeStyle(ws, 2, 3, 2, 7, (cell) => {
    cell.font = { name: 'Microsoft JhengHei', size: 16, bold: true, color: { argb: COLORS.TEXT_MAIN } };
    cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  });
  ws.getRow(2).height = 28;

  // Row 3: Metadata (C3:G3) - Invoice No & Date
  ws.mergeCells('C3:G3');
  const metaCell1 = ws.getCell('C3');
  metaCell1.value = '請款單號：{{INVOICE_NO | INV-202610-01}}  ｜  請款日期：{{INVOICE_DATE | 2026-09-30}}';
  applyCellRangeStyle(ws, 3, 3, 3, 7, (cell) => {
    cell.font = { name: 'Microsoft JhengHei', size: 9.5, color: { argb: COLORS.TEXT_MUTED } };
    cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    cell.border = BORDER_STYLE;
  });
  ws.getRow(3).height = 20;

  // Row 4: Metadata (C4:G4) - Contract Name & Due Date
  ws.mergeCells('C4:G4');
  const metaCell2 = ws.getCell('C4');
  metaCell2.value = '專案名稱：{{CONTRACT_NAME | 企業 AI 智慧客服整合開發案}}  ｜  付款截止：{{DUE_DATE | 2026-10-15}}';
  applyCellRangeStyle(ws, 4, 3, 4, 7, (cell) => {
    cell.font = { name: 'Microsoft JhengHei', size: 9.5, color: { argb: COLORS.TEXT_MUTED } };
    cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    cell.border = BORDER_STYLE;
  });
  ws.getRow(4).height = 20;

  // Row 5: Blank separator
  ws.getRow(5).height = 12;

  // Row 6: Section Headers (A6:D6 & E6:G6)
  ws.mergeCells('A6:D6');
  ws.getCell('A6').value = '【 受款單位資訊 (Payee) 】';
  applyCellRangeStyle(ws, 6, 1, 6, 4, (cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.SECTION_HEADER_BG } };
    cell.font = { name: 'Microsoft JhengHei', size: 9.5, bold: true, color: { argb: COLORS.SECTION_HEADER_TEXT } };
    cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    cell.border = BORDER_STYLE;
  });

  ws.mergeCells('E6:G6');
  ws.getCell('E6').value = '【 買受機構資料 (Payer) 】';
  applyCellRangeStyle(ws, 6, 5, 6, 7, (cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.SECTION_HEADER_BG } };
    cell.font = { name: 'Microsoft JhengHei', size: 9.5, bold: true, color: { argb: COLORS.SECTION_HEADER_TEXT } };
    cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    cell.border = BORDER_STYLE;
  });
  ws.getRow(6).height = 22;

  // Rows 7-10: Vendor & Client Information Fields
  const partyRows = [
    {
      r: 7,
      left: '受款公司：{{VENDOR_NAME | 智創雲端系統整合股份有限公司}}',
      right: '客戶名稱：{{CLIENT_COMPANY | 台灣智慧創新金融科技股份有限公司}}',
    },
    {
      r: 8,
      left: '統一編號：{{VENDOR_TAX_ID | 54321098}}',
      right: '統一編號：{{CLIENT_TAX_ID | 87654321}}',
    },
    {
      r: 9,
      left: '聯絡電話：{{VENDOR_PHONE | (02) 8765-4321}}',
      right: '聯絡窗口：{{CLIENT_CONTACT | 林經理 (0933-221-100)}}',
    },
    {
      r: 10,
      left: '通訊地址：{{VENDOR_ADDRESS | 台北市內湖區科技路一段88號}}',
      right: '通訊地址：{{CLIENT_ADDRESS | 台北市信義區松仁路100號}}',
    },
  ];

  partyRows.forEach(item => {
    ws.getRow(item.r).height = 20;

    ws.mergeCells(`A${item.r}:D${item.r}`);
    ws.getCell(`A${item.r}`).value = item.left;
    applyCellRangeStyle(ws, item.r, 1, item.r, 4, (cell) => {
      cell.font = { name: 'Microsoft JhengHei', size: 9, color: { argb: COLORS.TEXT_MAIN } };
      cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
      cell.border = BORDER_STYLE;
    });

    ws.mergeCells(`E${item.r}:G${item.r}`);
    ws.getCell(`E${item.r}`).value = item.right;
    applyCellRangeStyle(ws, item.r, 5, item.r, 7, (cell) => {
      cell.font = { name: 'Microsoft JhengHei', size: 9, color: { argb: COLORS.TEXT_MAIN } };
      cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
      cell.border = BORDER_STYLE;
    });
  });

  // Row 11: Blank separator
  ws.getRow(11).height = 12;

  // Row 12: Table Header
  const headers = [
    { col: 'A', text: '序號', align: 'center' },
    { col: 'B', text: '請款項目 / 階段', align: 'center' },
    { col: 'C', text: '工作內容說明與交付成果', align: 'center' },
    { col: 'D', text: '數量', align: 'center' },
    { col: 'E', text: '單位', align: 'center' },
    { col: 'F', text: '單價 (NTD)', align: 'center' },
    { col: 'G', text: '金額 (NTD)', align: 'center' },
  ];

  headers.forEach(h => {
    const c = ws.getCell(`${h.col}12`);
    c.value = h.text;
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.TABLE_HEADER_BG } };
    c.font = { name: 'Microsoft JhengHei', size: 9.5, bold: true, color: { argb: COLORS.TABLE_HEADER_TEXT } };
    c.alignment = { vertical: 'middle', horizontal: h.align };
    c.border = BORDER_HEADER;
  });
  ws.getRow(12).height = 24;

  // Rows 13-15: Stage Items
  const items = [
    {
      r: 13,
      no: 1,
      name: '{{STAGE_1_NAME | 系統規格分析與原型設計完成}}',
      desc: '{{STAGE_1_DESC | 完成需求訪談規格書 (SRS) 與 Figma UI/UX 互動原型驗收}}',
      qty: '{{STAGE_1_QTY | 1}}',
      unit: '{{STAGE_1_UNIT | 期}}',
      price: '{{STAGE_1_PRICE | 120000}}',
    },
    {
      r: 14,
      no: 2,
      name: '{{STAGE_2_NAME | 核心 AI 對話引擎開發與 API 串接}}',
      desc: '{{STAGE_2_DESC | 整合企業知識庫 RAG 模型、智慧意圖識別與中台介接完成}}',
      qty: '{{STAGE_2_QTY | 1}}',
      unit: '{{STAGE_2_UNIT | 期}}',
      price: '{{STAGE_2_PRICE | 180000}}',
    },
    {
      r: 15,
      no: 3,
      name: '{{STAGE_3_NAME | 雲端伺服器建置與環境部署}}',
      desc: '{{STAGE_3_DESC | 生產環境高可用叢集架設、SSL 憑證與監控警報配置}}',
      qty: '{{STAGE_3_QTY | 1}}',
      unit: '{{STAGE_3_UNIT | 式}}',
      price: '{{STAGE_3_PRICE | 60000}}',
    },
  ];

  items.forEach(it => {
    ws.getRow(it.r).height = 22;

    const cA = ws.getCell(`A${it.r}`);
    cA.value = it.no;
    cA.alignment = { vertical: 'middle', horizontal: 'center' };
    cA.font = { name: 'Microsoft JhengHei', size: 9.5, color: { argb: COLORS.TEXT_MAIN } };
    cA.border = BORDER_STYLE;

    const cB = ws.getCell(`B${it.r}`);
    cB.value = it.name;
    cB.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    cB.font = { name: 'Microsoft JhengHei', size: 9.5, color: { argb: COLORS.TEXT_MAIN } };
    cB.border = BORDER_STYLE;

    const cC = ws.getCell(`C${it.r}`);
    cC.value = it.desc;
    cC.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    cC.font = { name: 'Microsoft JhengHei', size: 9, color: { argb: COLORS.TEXT_MUTED } };
    cC.border = BORDER_STYLE;

    const cD = ws.getCell(`D${it.r}`);
    cD.value = it.qty;
    cD.alignment = { vertical: 'middle', horizontal: 'center' };
    cD.font = { name: 'Microsoft JhengHei', size: 9.5, color: { argb: COLORS.TEXT_MAIN } };
    cD.border = BORDER_STYLE;

    const cE = ws.getCell(`E${it.r}`);
    cE.value = it.unit;
    cE.alignment = { vertical: 'middle', horizontal: 'center' };
    cE.font = { name: 'Microsoft JhengHei', size: 9.5, color: { argb: COLORS.TEXT_MAIN } };
    cE.border = BORDER_STYLE;

    const cF = ws.getCell(`F${it.r}`);
    cF.value = it.price;
    cF.alignment = { vertical: 'middle', horizontal: 'right' };
    cF.font = { name: 'Microsoft JhengHei', size: 9.5, color: { argb: COLORS.TEXT_MAIN } };
    cF.numFmt = '"NT$ "#,##0';
    cF.border = BORDER_STYLE;

    const cG = ws.getCell(`G${it.r}`);
    cG.value = { formula: `D${it.r}*F${it.r}` };
    cG.alignment = { vertical: 'middle', horizontal: 'right' };
    cG.font = { name: 'Microsoft JhengHei', size: 9.5, color: { argb: COLORS.TEXT_MAIN } };
    cG.numFmt = '"NT$ "#,##0';
    cG.border = BORDER_STYLE;
  });

  // Row 16: Subtotal
  ws.mergeCells('A16:F16');
  ws.getCell('A16').value = '請款金額合計 (未稅 Subtotal)';
  applyCellRangeStyle(ws, 16, 1, 16, 6, (cell) => {
    cell.font = { name: 'Microsoft JhengHei', size: 9.5, color: { argb: COLORS.TEXT_MAIN } };
    cell.alignment = { vertical: 'middle', horizontal: 'right' };
    cell.border = BORDER_STYLE;
  });
  const cG16 = ws.getCell('G16');
  cG16.value = { formula: 'SUM(G13:G15)' };
  cG16.font = { name: 'Microsoft JhengHei', size: 9.5, color: { argb: COLORS.TEXT_MAIN } };
  cG16.alignment = { vertical: 'middle', horizontal: 'right' };
  cG16.numFmt = '"NT$ "#,##0';
  cG16.border = BORDER_STYLE;
  ws.getRow(16).height = 22;

  // Row 17: VAT 5%
  ws.mergeCells('A17:F17');
  ws.getCell('A17').value = '加值型營業稅 (VAT 5%)';
  applyCellRangeStyle(ws, 17, 1, 17, 6, (cell) => {
    cell.font = { name: 'Microsoft JhengHei', size: 9.5, color: { argb: COLORS.TEXT_MAIN } };
    cell.alignment = { vertical: 'middle', horizontal: 'right' };
    cell.border = BORDER_STYLE;
  });
  const cG17 = ws.getCell('G17');
  cG17.value = { formula: 'ROUND(G16*0.05, 0)' };
  cG17.font = { name: 'Microsoft JhengHei', size: 9.5, color: { argb: COLORS.TEXT_MAIN } };
  cG17.alignment = { vertical: 'middle', horizontal: 'right' };
  cG17.numFmt = '"NT$ "#,##0';
  cG17.border = BORDER_STYLE;
  ws.getRow(17).height = 22;

  // Row 18: Grand Total
  ws.mergeCells('A18:F18');
  ws.getCell('A18').value = '本期應請款總額 (含稅 Total Due NTD)';
  applyCellRangeStyle(ws, 18, 1, 18, 6, (cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.TOTAL_BG } };
    cell.font = { name: 'Microsoft JhengHei', size: 10, bold: true, color: { argb: COLORS.TOTAL_TEXT } };
    cell.alignment = { vertical: 'middle', horizontal: 'right' };
    cell.border = {
      top: { style: 'thin', color: { argb: COLORS.BORDER_GREY } },
      left: { style: 'thin', color: { argb: COLORS.BORDER_GREY } },
      bottom: { style: 'double', color: { argb: COLORS.TEXT_MAIN } },
      right: { style: 'thin', color: { argb: COLORS.BORDER_GREY } },
    };
  });
  const cG18 = ws.getCell('G18');
  cG18.value = { formula: 'G16+G17' };
  cG18.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.TOTAL_BG } };
  cG18.font = { name: 'Microsoft JhengHei', size: 10.5, bold: true, color: { argb: 'FF1E3A8A' } };
  cG18.alignment = { vertical: 'middle', horizontal: 'right' };
  cG18.numFmt = '"NT$ "#,##0';
  cG18.border = {
    top: { style: 'thin', color: { argb: COLORS.BORDER_GREY } },
    left: { style: 'thin', color: { argb: COLORS.BORDER_GREY } },
    bottom: { style: 'double', color: { argb: COLORS.TEXT_MAIN } },
    right: { style: 'thin', color: { argb: COLORS.BORDER_GREY } },
  };
  ws.getRow(18).height = 24;

  // Row 19: Blank separator
  ws.getRow(19).height = 12;

  // Row 20: Bank Section Header
  ws.mergeCells('A20:G20');
  ws.getCell('A20').value = '【 指定金融匯款帳戶資訊 】';
  applyCellRangeStyle(ws, 20, 1, 20, 7, (cell) => {
    cell.font = { name: 'Microsoft JhengHei', size: 9.5, bold: true, color: { argb: COLORS.TEXT_MAIN } };
    cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    cell.border = BORDER_STYLE;
  });
  ws.getRow(20).height = 22;

  // Rows 21-22: Bank Details
  ws.mergeCells('A21:G21');
  ws.getCell('A21').value = '匯款銀行：{{BANK_NAME | 國泰世華商業銀行 (013)}}        分行：{{BANK_BRANCH | 敦南分行}}';
  applyCellRangeStyle(ws, 21, 1, 21, 7, (cell) => {
    cell.font = { name: 'Microsoft JhengHei', size: 9, color: { argb: COLORS.TEXT_MAIN } };
    cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    cell.border = BORDER_STYLE;
  });
  ws.getRow(21).height = 20;

  ws.mergeCells('A22:G22');
  ws.getCell('A22').value = '銀行帳號：{{BANK_ACCOUNT | 028-50-6008899}}        受款戶名：{{BANK_ACCOUNT_NAME | 智創雲端系統整合股份有限公司}}';
  applyCellRangeStyle(ws, 22, 1, 22, 7, (cell) => {
    cell.font = { name: 'Microsoft JhengHei', size: 9, color: { argb: COLORS.TEXT_MAIN } };
    cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    cell.border = BORDER_STYLE;
  });
  ws.getRow(22).height = 20;

  // Row 23: Blank separator
  ws.getRow(23).height = 12;

  // Row 24: Terms Section Header
  ws.mergeCells('A24:G24');
  ws.getCell('A24').value = '【 請款備註與付款約定事項 】';
  applyCellRangeStyle(ws, 24, 1, 24, 7, (cell) => {
    cell.font = { name: 'Microsoft JhengHei', size: 9.5, bold: true, color: { argb: COLORS.TEXT_MAIN } };
    cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    cell.border = BORDER_STYLE;
  });
  ws.getRow(24).height = 22;

  // Rows 25-26: Terms details
  const terms = [
    { r: 25, text: '1. {{PAYMENT_NOTE_1 | 本請款單依合約第二條請款進度開立，請於付款截止日前安排撥款，手續費由受款方吸收。}}' },
    { r: 26, text: '2. {{PAYMENT_NOTE_2 | 款項匯出後敬請回傳匯款單據或末五碼至 finance@zhichuang.com，本公司將於確認入帳後開立統一發票。}}' },
  ];

  terms.forEach(t => {
    ws.getRow(t.r).height = 20;
    ws.mergeCells(`A${t.r}:G${t.r}`);
    ws.getCell(`A${t.r}`).value = t.text;
    applyCellRangeStyle(ws, t.r, 1, t.r, 7, (cell) => {
      cell.font = { name: 'Microsoft JhengHei', size: 9, color: { argb: COLORS.TEXT_MUTED } };
      cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
      cell.border = BORDER_STYLE;
    });
  });

  // Row 27: Blank separator
  ws.getRow(27).height = 12;

  // Rows 28-30: Signature box
  ws.mergeCells('A28:C28');
  ws.getCell('A28').value = '請款單位主管用印蓋章';
  applyCellRangeStyle(ws, 28, 1, 28, 3, (cell) => {
    cell.font = { name: 'Microsoft JhengHei', size: 9, color: { argb: COLORS.TEXT_MUTED } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = BORDER_STYLE;
  });

  ws.getCell('D28').border = BORDER_STYLE;

  ws.mergeCells('E28:G28');
  ws.getCell('E28').value = '買受機構權責主管覆核簽章';
  applyCellRangeStyle(ws, 28, 5, 28, 7, (cell) => {
    cell.font = { name: 'Microsoft JhengHei', size: 9, color: { argb: COLORS.TEXT_MUTED } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = BORDER_STYLE;
  });
  ws.getRow(28).height = 20;

  ws.mergeCells('A29:C31');
  ws.getCell('A29').value = '\n（請加蓋公司用印或發票專用章）\n日期：   年   月   日';
  applyCellRangeStyle(ws, 29, 1, 31, 3, (cell) => {
    cell.font = { name: 'Microsoft JhengHei', size: 8.5, color: { argb: COLORS.TEXT_LIGHT } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = BORDER_STYLE;
  });

  for (let r = 29; r <= 31; r++) {
    ws.getCell(`D${r}`).border = BORDER_STYLE;
    ws.getRow(r).height = 18;
  }

  ws.mergeCells('E29:G31');
  ws.getCell('E29').value = '\n（請覆核人簽名並加蓋權責用印）\n日期：   年   月   日';
  applyCellRangeStyle(ws, 29, 5, 31, 7, (cell) => {
    cell.font = { name: 'Microsoft JhengHei', size: 8.5, color: { argb: COLORS.TEXT_LIGHT } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = BORDER_STYLE;
  });

  await wb.xlsx.writeFile('public/商業請款單範本.xlsx');
  console.log('Saved 100% matched public/商業請款單範本.xlsx');
}

async function main() {
  await buildQuotationTemplate();
  await buildInvoiceTemplate();
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
