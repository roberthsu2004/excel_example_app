import React, { useState, useEffect, useMemo } from 'react';
import {
  FileSpreadsheet,
  Download,
  RotateCcw,
  Eraser,
  Search,
  CheckCircle2,
  AlertCircle,
  FileText,
  Sparkles,
  Layers,
  Table,
  Eye,
  Info,
  ChevronRight,
  Calculator,
  Image as ImageIcon,
  Copy,
  Check,
  X,
  RefreshCw,
  Calendar,
  ListFilter,
  ChevronDown,
  Plus,
  Trash2,
  ListOrdered,
  Building2,
  CreditCard,
  FileSignature,
} from 'lucide-react';
import { PlaceholderVariable, TemplateMeta, ToastNotification, QuotationItem } from './types/template';
import {
  loadWorkbook,
  scanWorkbookPlaceholders,
  exportFilledWorkbook,
  generateFallbackTemplate,
  formatKeyToChinese,
  extractQuotationItems,
  createNextItem,
} from './utils/excelEngine';

export default function App() {
  // State
  const [templateMeta, setTemplateMeta] = useState<TemplateMeta | null>(null);
  const [variables, setVariables] = useState<PlaceholderVariable[]>([]);
  const [values, setValues] = useState<Record<string, string>>({});
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedZone, setSelectedZone] = useState<'all' | 'header' | 'items' | 'footer'>('all');
  const [viewMode, setViewMode] = useState<'form' | 'items' | 'table'>('form');
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [isExporting, setIsExporting] = useState<boolean>(false);
  const [activePreset, setActivePreset] = useState<'quote' | 'invoice'>('quote');
  const [copiedJson, setCopiedJson] = useState(false);
  const [toasts, setToasts] = useState<ToastNotification[]>([]);
  const [hasCustomLogo, setHasCustomLogo] = useState<boolean>(false);
  const [embedLogoImage, setEmbedLogoImage] = useState<boolean>(false);
  const [tableExportMode, setTableExportMode] = useState<'auto-shrink' | 'keep-template-rows'>('auto-shrink');

  // Toast helper
  const addToast = (type: ToastNotification['type'], title: string, message: string) => {
    const id = Math.random().toString(36).substring(2, 9);
    setToasts((prev) => [...prev, { id, type, title, message }]);
    setTimeout(() => {
      setToasts((prev) => prev.filter((t) => t.id !== id));
    }, 4500);
  };

  const removeToast = (id: string) => {
    setToasts((prev) => prev.filter((t) => t.id !== id));
  };

  // Process ArrayBuffer to extract placeholders and metadata
  const processWorkbookBuffer = async (buffer: ArrayBuffer, fileName: string) => {
    try {
      setIsLoading(true);
      const workbook = await loadWorkbook(buffer);
      const scanResult = scanWorkbookPlaceholders(workbook);

      const sheetNames = workbook.worksheets.map((ws) => ws.name);

      setTemplateMeta({
        fileName,
        fileSize: buffer.byteLength,
        sheetNames,
        totalCellsScanned: scanResult.totalCellsScanned,
        formulaCount: scanResult.formulaCount,
        imageCount: scanResult.imageCount,
        hasHeaderFooter: scanResult.hasHeaderFooter,
        rawBuffer: buffer,
      });

      setVariables(scanResult.variables);

      // Pre-fill values with defaults
      const initialValues: Record<string, string> = {};
      scanResult.variables.forEach((v) => {
        initialValues[v.key] = v.defaultValue;
      });
      setValues(initialValues);

      addToast(
        'success',
        '範本解析成功',
        `已成功載入「${fileName}」，偵測到 ${scanResult.variables.length} 個佔位符變數與 ${scanResult.formulaCount} 個動態計算公式。`
      );
    } catch (err: any) {
      console.error('Failed to parse workbook:', err);
      addToast('error', '檔案解析失敗', err.message || '無法正確讀取該 Excel 檔案，請確認格式為標準 .xlsx。');
    } finally {
      setIsLoading(false);
    }
  };

  // Load standard preset template
  const loadPresetTemplate = async (type: 'quote' | 'invoice') => {
    setActivePreset(type);
    setIsLoading(true);
    const fileName = type === 'quote' ? '商務報價單範本.xlsx' : '商業請款單範本.xlsx';

    try {
      const response = await fetch(`/${encodeURIComponent(fileName)}?t=${Date.now()}`);
      if (!response.ok) {
        throw new Error(`無法載入伺服器範本 (${response.status})，即將啟動備援範本生成器。`);
      }
      const buffer = await response.arrayBuffer();
      await processWorkbookBuffer(buffer, fileName);
    } catch (err: any) {
      console.warn('Fetch failed, generating in-memory fallback template:', err);
      // Fallback: Generate pristine in-memory template
      const fallbackBuf = await generateFallbackTemplate(type);
      await processWorkbookBuffer(fallbackBuf, fileName);
      addToast(
        'info',
        '已載入內建備援範本',
        `已自動為您初始化標準「${type === 'quote' ? '商務報價單' : '商業請款單'}」範本架構。`
      );
    }
  };

  // On initial mount: Check for logo.png and load default quote template
  useEffect(() => {
    fetch('/logo.png', { method: 'HEAD' })
      .then((res) => {
        if (res.ok) setHasCustomLogo(true);
      })
      .catch(() => {});
    loadPresetTemplate('quote');
  }, []);

  // Reset to default values
  const handleResetToDefaults = () => {
    const defaultVals: Record<string, string> = {};
    variables.forEach((v) => {
      defaultVals[v.key] = v.defaultValue;
    });
    setValues(defaultVals);
    addToast('info', '已重設回預設值', '所有欄位已還原為範本最初設定的預設內容。');
  };

  // Clear all values
  const handleClearAll = () => {
    const emptyVals: Record<string, string> = {};
    variables.forEach((v) => {
      emptyVals[v.key] = '';
    });
    setValues(emptyVals);
    addToast('info', '已清空所有欄位', '目前所有輸入框已全部清空，您可以重新填入所需數值。');
  };

  // Export and download
  const handleExport = async () => {
    if (!templateMeta) {
      addToast('warning', '無法匯出', '請先載入或上傳 Excel 範本檔案。');
      return;
    }

    try {
      setIsExporting(true);
      await exportFilledWorkbook(templateMeta.rawBuffer, values, templateMeta.fileName, {
        embedLogo: embedLogoImage,
        tableMode: tableExportMode,
      });
      addToast(
        'success',
        '匯出成功',
        `已成功填入數值並觸發下載！原範本中的 Logo 圖片、格式、框線與計算公式已完整保留。`
      );
    } catch (err: any) {
      console.error('Export error:', err);
      addToast('error', '匯出失敗', err.message || '在產生 Excel 檔案時發生錯誤。');
    } finally {
      setIsExporting(false);
    }
  };

  // Copy current values as JSON
  const handleCopyJson = () => {
    const jsonStr = JSON.stringify(values, null, 2);
    navigator.clipboard.writeText(jsonStr);
    setCopiedJson(true);
    addToast('success', '複製成功', '已將所有欄位變數與數值複製為 JSON 格式。');
    setTimeout(() => setCopiedJson(false), 2000);
  };

  // Date formatting and quick setter helpers
  const normalizeDateValue = (val: string): string => {
    if (!val) return '';
    const trimmed = val.trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) return trimmed;
    if (/^\d{4}\/\d{1,2}\/\d{1,2}$/.test(trimmed)) {
      const parts = trimmed.split('/');
      return `${parts[0]}-${parts[1].padStart(2, '0')}-${parts[2].padStart(2, '0')}`;
    }
    const d = new Date(trimmed);
    if (!isNaN(d.getTime())) {
      const year = d.getFullYear();
      const month = String(d.getMonth() + 1).padStart(2, '0');
      const day = String(d.getDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    }
    return '';
  };

  const setQuickDate = (key: string, offsetDays: number) => {
    const d = new Date();
    d.setDate(d.getDate() + offsetDays);
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    const dateStr = `${year}-${month}-${day}`;
    setValues((prev) => ({
      ...prev,
      [key]: dateStr,
    }));
  };

  const setEndOfMonth = (key: string) => {
    const d = new Date();
    const end = new Date(d.getFullYear(), d.getMonth() + 1, 0);
    const year = end.getFullYear();
    const month = String(end.getMonth() + 1).padStart(2, '0');
    const day = String(end.getDate()).padStart(2, '0');
    const dateStr = `${year}-${month}-${day}`;
    setValues((prev) => ({
      ...prev,
      [key]: dateStr,
    }));
  };

  // Zone ordering weights for natural top-to-bottom layout
  const HEADER_KEY_ORDER: Record<string, number> = {
    公司商標名稱: 1,
    COMPANY_LOGO: 1,
    報價單號: 2,
    QUOTATION_NO: 2,
    請款單號: 2,
    INVOICE_NO: 2,
    報價日期: 3,
    QUOTATION_DATE: 3,
    請款日期: 3,
    INVOICE_DATE: 3,
    有效天數: 4,
    VALID_DAYS: 4,
    付款截止日: 4,
    DUE_DATE: 4,
    專案名稱: 5,
    PROJECT_NAME: 5,
    專案合約名稱: 5,
    CONTRACT_NAME: 5,
    報價公司名稱: 6,
    PROVIDER_NAME: 6,
    受款公司名稱: 6,
    VENDOR_NAME: 6,
    客戶公司名稱: 7,
    CLIENT_NAME: 7,
    CLIENT_COMPANY: 7,
    報價公司統編: 8,
    PROVIDER_TAX_ID: 8,
    受款公司統編: 8,
    VENDOR_TAX_ID: 8,
    客戶統編: 9,
    CLIENT_TAX_ID: 9,
    報價公司電話: 10,
    PROVIDER_PHONE: 10,
    受款公司電話: 10,
    VENDOR_PHONE: 10,
    客戶聯絡窗口: 11,
    CLIENT_CONTACT: 11,
    買受聯絡窗口: 11,
    報價公司地址: 12,
    PROVIDER_ADDRESS: 12,
    受款公司通訊地址: 12,
    VENDOR_ADDRESS: 12,
    客戶通訊地址: 13,
    CLIENT_ADDRESS: 13,
    買受送單地址: 13,
  };

  const FOOTER_KEY_ORDER: Record<string, number> = {
    受款銀行名稱: 1,
    BANK_NAME: 1,
    受款銀行分行: 2,
    BANK_BRANCH: 2,
    受款銀行帳號: 3,
    BANK_ACCOUNT: 3,
    受款戶名: 4,
    BANK_ACCOUNT_NAME: 4,
    付款條件說明: 5,
    PAYMENT_TERMS: 5,
    匯款專戶資訊: 6,
    有效期限條款: 7,
    VALID_TERMS: 7,
    款項備註說明_1: 8,
    PAYMENT_NOTE_1: 8,
    款項備註說明_2: 9,
    PAYMENT_NOTE_2: 9,
  };

  const isItemVar = (v: PlaceholderVariable): boolean => {
    const k = v.key.toUpperCase();
    return (
      v.category === '明細項目' ||
      k.includes('ITEM') ||
      k.includes('STAGE') ||
      k.includes('品項') ||
      k.includes('請款階段')
    );
  };

  const isFooterVar = (v: PlaceholderVariable): boolean => {
    if (isItemVar(v)) return false;
    const k = v.key.toUpperCase();
    return (
      k.includes('TERMS') ||
      k.includes('BANK') ||
      k.includes('PAYMENT') ||
      k.includes('NOTE') ||
      k.includes('MEMO') ||
      k.includes('條款') ||
      k.includes('備註') ||
      k.includes('帳戶') ||
      k.includes('銀行') ||
      k.includes('戶名') ||
      k.includes('帳號') ||
      k.includes('說明_') ||
      k.includes('款項')
    );
  };

  const isHeaderVar = (v: PlaceholderVariable): boolean => {
    return !isItemVar(v) && !isFooterVar(v);
  };

  const matchesQuery = (v: PlaceholderVariable) => {
    if (!searchQuery.trim()) return true;
    const q = searchQuery.trim().toLowerCase();
    const zhLabel = formatKeyToChinese(v.key).toLowerCase();
    return (
      v.key.toLowerCase().includes(q) ||
      zhLabel.includes(q) ||
      (values[v.key] || '').toLowerCase().includes(q) ||
      v.defaultValue.toLowerCase().includes(q) ||
      v.occurrences.some((occ) => occ.cellAddress.toLowerCase().includes(q) || occ.sheetName.toLowerCase().includes(q))
    );
  };

  // Header Zone variables sorted top-to-bottom
  const headerVariables = useMemo(() => {
    return variables
      .filter(isHeaderVar)
      .sort((a, b) => {
        const orderA = HEADER_KEY_ORDER[a.key] || HEADER_KEY_ORDER[formatKeyToChinese(a.key)] || 99;
        const orderB = HEADER_KEY_ORDER[b.key] || HEADER_KEY_ORDER[formatKeyToChinese(b.key)] || 99;
        return orderA - orderB;
      });
  }, [variables]);

  // Footer Zone variables sorted top-to-bottom
  const footerVariables = useMemo(() => {
    return variables
      .filter(isFooterVar)
      .sort((a, b) => {
        const orderA = FOOTER_KEY_ORDER[a.key] || FOOTER_KEY_ORDER[formatKeyToChinese(a.key)] || 99;
        const orderB = FOOTER_KEY_ORDER[b.key] || FOOTER_KEY_ORDER[formatKeyToChinese(b.key)] || 99;
        return orderA - orderB;
      });
  }, [variables]);

  const filteredHeaderVars = useMemo(() => headerVariables.filter(matchesQuery), [headerVariables, searchQuery, values]);
  const filteredFooterVars = useMemo(() => footerVariables.filter(matchesQuery), [footerVariables, searchQuery, values]);
  const filteredVariables = useMemo(() => variables.filter(matchesQuery), [variables, searchQuery, values]);

  // Statistics
  const stats = useMemo(() => {
    let modifiedCount = 0;
    let emptyCount = 0;
    variables.forEach((v) => {
      const current = values[v.key] ?? '';
      if (current !== v.defaultValue) modifiedCount++;
      if (current.trim() === '') emptyCount++;
    });

    return {
      total: variables.length,
      modified: modifiedCount,
      empty: emptyCount,
    };
  }, [variables, values]);

  // Extract quotation items for Table Mode
  const quotationItems = useMemo(() => {
    return extractQuotationItems(variables, values);
  }, [variables, values]);

  // Live financial totals calculated from table items
  const { totalSubtotal, totalVat, grandTotal } = useMemo(() => {
    let subtotal = 0;
    quotationItems.forEach((item) => {
      const q = parseFloat(String(values[item.qtyKey] !== undefined ? values[item.qtyKey] : item.qty)) || 0;
      const p = parseFloat(String(values[item.priceKey] !== undefined ? values[item.priceKey] : item.price)) || 0;
      subtotal += q * p;
    });
    const vat = Math.round(subtotal * 0.05);
    const grand = subtotal + vat;
    return { totalSubtotal: subtotal, totalVat: vat, grandTotal: grand };
  }, [quotationItems, values]);

  // Add new item row handler
  const handleAddItem = () => {
    const { newItem, newVariables, newValues } = createNextItem(quotationItems, variables);
    setVariables((prev) => [...prev, ...newVariables]);
    setValues((prev) => ({
      ...prev,
      ...newValues,
    }));
    addToast('success', '已新增品項', `已加入「${newItem.name}」，一列代表一個品項。`);
  };

  // Delete item row handler
  const handleDeleteItem = (item: QuotationItem) => {
    setValues((prev) => {
      const next = { ...prev };
      delete next[item.nameKey];
      delete next[item.descKey];
      delete next[item.qtyKey];
      delete next[item.unitKey];
      delete next[item.priceKey];
      return next;
    });
    setVariables((prev) =>
      prev.filter(
        (v) =>
          v.key !== item.nameKey &&
          v.key !== item.descKey &&
          v.key !== item.qtyKey &&
          v.key !== item.unitKey &&
          v.key !== item.priceKey
      )
    );
    addToast('info', '已刪除品項', `已移除品項「${item.name || item.rowNumber}」。`);
  };

  // Helper to render individual input card for Header and Footer zones
  const renderFieldCard = (v: PlaceholderVariable) => {
    const zhLabel = formatKeyToChinese(v.key);
    const currentValue = values[v.key] ?? '';
    const isModified = currentValue !== v.defaultValue;
    const isEmpty = currentValue.trim() === '';

    return (
      <div
        key={v.key}
        className={`p-4 rounded-xl border transition-all ${
          isModified
            ? 'bg-blue-50/30 border-blue-300 ring-1 ring-blue-500/10 shadow-xs'
            : isEmpty
            ? 'bg-amber-50/20 border-amber-200/80 shadow-2xs'
            : 'bg-white border-slate-200/90 shadow-2xs hover:border-slate-300'
        }`}
      >
        <div className="flex items-start justify-between gap-2 mb-2">
          <div>
            <label className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
              <span>{zhLabel}</span>
              {isModified && (
                <span className="text-[10px] bg-blue-100 text-blue-700 font-semibold px-1.5 py-0.2 rounded">
                  已編輯
                </span>
              )}
            </label>
            {zhLabel !== v.key && (
              <span className="text-[10px] font-mono text-slate-400 block mt-0.5">
                {`{{${v.key}}}`}
              </span>
            )}
          </div>

          <div className="flex items-center gap-1 shrink-0">
            {v.occurrences.map((occ) => (
              <span
                key={occ.cellAddress}
                className="text-[10px] font-mono font-medium px-1.5 py-0.5 rounded bg-slate-100 text-slate-600 border border-slate-200"
                title={`${occ.sheetName} 儲存格 ${occ.cellAddress}`}
              >
                {occ.cellAddress}
              </span>
            ))}
          </div>
        </div>

        {/* Input Control */}
        <div className="space-y-1.5">
          {v.isDate ? (
            <div className="space-y-1.5">
              <div className="relative">
                <input
                  type="date"
                  value={normalizeDateValue(currentValue)}
                  onChange={(e) =>
                    setValues((prev) => ({
                      ...prev,
                      [v.key]: e.target.value,
                    }))
                  }
                  className="w-full text-xs px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 font-mono font-medium text-slate-800 transition-all cursor-pointer"
                />
              </div>
              <div className="flex items-center gap-1 flex-wrap pt-0.5">
                <span className="text-[10px] text-slate-400 font-medium">快速設定:</span>
                <button
                  type="button"
                  onClick={() => setQuickDate(v.key, 0)}
                  className="text-[10px] font-medium px-1.5 py-0.5 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded transition-colors cursor-pointer"
                >
                  今天
                </button>
                <button
                  type="button"
                  onClick={() => setQuickDate(v.key, 7)}
                  className="text-[10px] font-medium px-1.5 py-0.5 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded transition-colors cursor-pointer"
                >
                  +7天
                </button>
                <button
                  type="button"
                  onClick={() => setQuickDate(v.key, 14)}
                  className="text-[10px] font-medium px-1.5 py-0.5 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded transition-colors cursor-pointer"
                >
                  +14天
                </button>
                <button
                  type="button"
                  onClick={() => setQuickDate(v.key, 30)}
                  className="text-[10px] font-medium px-1.5 py-0.5 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded transition-colors cursor-pointer"
                >
                  +30天
                </button>
                <button
                  type="button"
                  onClick={() => setEndOfMonth(v.key)}
                  className="text-[10px] font-medium px-1.5 py-0.5 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded transition-colors cursor-pointer"
                >
                  月底
                </button>
              </div>
            </div>
          ) : v.dropdownOptions && v.dropdownOptions.length > 0 ? (
            <div className="flex items-center gap-1.5">
              <select
                value={
                  v.dropdownOptions.includes(currentValue)
                    ? currentValue
                    : '__custom__'
                }
                onChange={(e) => {
                  if (e.target.value !== '__custom__') {
                    setValues((prev) => ({
                      ...prev,
                      [v.key]: e.target.value,
                    }));
                  }
                }}
                className="text-xs px-2.5 py-2 bg-white border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 font-medium text-slate-800 transition-all cursor-pointer shrink-0"
              >
                <option value="" disabled>-- 預設選項 --</option>
                {v.dropdownOptions.map((opt) => (
                  <option key={opt} value={opt}>
                    {opt}
                  </option>
                ))}
                <option value="__custom__">✏️ 自訂</option>
              </select>
              <input
                type="text"
                value={currentValue}
                onChange={(e) =>
                  setValues((prev) => ({
                    ...prev,
                    [v.key]: e.target.value,
                  }))
                }
                placeholder={v.defaultValue || '請輸入內容'}
                className="flex-1 text-xs px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 font-medium text-slate-800 transition-all"
              />
            </div>
          ) : v.isLongText ? (
            <textarea
              rows={2}
              value={currentValue}
              onChange={(e) =>
                setValues((prev) => ({
                  ...prev,
                  [v.key]: e.target.value,
                }))
              }
              placeholder={v.defaultValue || '請輸入多行說明內容'}
              className="w-full text-xs px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 text-slate-800 transition-all resize-y"
            />
          ) : (
            <input
              type="text"
              value={currentValue}
              onChange={(e) =>
                setValues((prev) => ({
                  ...prev,
                  [v.key]: e.target.value,
                }))
              }
              placeholder={v.defaultValue || '請輸入內容'}
              className="w-full text-xs px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 font-medium text-slate-800 transition-all"
            />
          )}

          {/* Bottom Default Hint */}
          <div className="flex items-center justify-between text-[11px] text-slate-400 pt-0.5">
            <span className="truncate max-w-[200px]" title={v.defaultValue}>
              預設: {v.defaultValue || '（無）'}
            </span>
            {isModified && (
              <button
                type="button"
                onClick={() =>
                  setValues((prev) => ({
                    ...prev,
                    [v.key]: v.defaultValue,
                  }))
                }
                className="text-blue-600 hover:text-blue-700 text-[10px] font-medium hover:underline cursor-pointer"
              >
                還原預設
              </button>
            )}
          </div>
        </div>
      </div>
    );
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 flex flex-col font-sans">
      {/* ================= Top Bar Contract (3 Zones) ================= */}
      <header className="sticky top-0 z-30 bg-white/95 backdrop-blur border-b border-slate-200 px-6 py-3.5 flex items-center justify-between shadow-xs">
        {/* Zone 1: Single text element wordmark + custom logo */}
        <div className="flex items-center gap-3">
          {hasCustomLogo ? (
            <img
              src={`/logo.png?t=${Date.now()}`}
              alt="Company Logo"
              className="h-9 w-auto max-h-9 max-w-[130px] object-contain rounded-md border border-slate-200 bg-white px-2 py-0.5 shadow-2xs"
              onError={() => setHasCustomLogo(false)}
            />
          ) : (
            <div className="w-9 h-9 rounded-lg bg-slate-900 text-white flex items-center justify-center shadow-xs">
              <FileSpreadsheet className="w-5 h-5 text-blue-400" />
            </div>
          )}
          <div>
            <h1 className="text-base font-bold tracking-tight text-slate-900 leading-tight">
              通用型 Excel 佔位符單據產生器
            </h1>
            <p className="text-xs text-slate-500 font-medium">
              免伺服器 · 零 API Key · 保留圖片與公式
            </p>
          </div>
        </div>

        {/* Zone 2: Navigation & Preset Loaders */}
        <nav className="flex items-center gap-1.5 p-1 bg-slate-100/90 rounded-lg border border-slate-200/80">
          <button
            type="button"
            onClick={() => loadPresetTemplate('quote')}
            className={`px-3.5 py-1.5 text-xs font-semibold rounded-md transition-all flex items-center gap-1.5 whitespace-nowrap cursor-pointer ${
              activePreset === 'quote'
                ? 'bg-white text-blue-700 shadow-xs border border-slate-200/60 font-bold'
                : 'text-slate-600 hover:text-slate-900 hover:bg-slate-200/50'
            }`}
          >
            <FileText className="w-3.5 h-3.5 text-blue-600" />
            📄 商務報價單範本
          </button>
          <button
            type="button"
            onClick={() => loadPresetTemplate('invoice')}
            className={`px-3.5 py-1.5 text-xs font-semibold rounded-md transition-all flex items-center gap-1.5 whitespace-nowrap cursor-pointer ${
              activePreset === 'invoice'
                ? 'bg-white text-blue-700 shadow-xs border border-slate-200/60 font-bold'
                : 'text-slate-600 hover:text-slate-900 hover:bg-slate-200/50'
            }`}
          >
            <FileSpreadsheet className="w-3.5 h-3.5 text-blue-600" />
            📑 商業請款單範本
          </button>
        </nav>

        {/* Zone 3: Primary Action */}
        <div className="flex items-center gap-2.5">
          <button
            type="button"
            onClick={handleExport}
            disabled={isLoading || isExporting || variables.length === 0}
            className="px-4 py-2 text-xs font-bold text-white bg-blue-600 hover:bg-blue-700 active:bg-blue-800 disabled:opacity-50 disabled:cursor-not-allowed rounded-lg shadow-sm transition-all flex items-center gap-2 whitespace-nowrap cursor-pointer"
          >
            {isExporting ? (
              <>
                <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                正在寫入儲存格...
              </>
            ) : (
              <>
                <Download className="w-3.5 h-3.5" />
                📥 匯出並下載 XLSX
              </>
            )}
          </button>
        </div>
      </header>

      {/* Main Workspace */}
      <main className="flex-1 max-w-7xl w-full mx-auto p-4 sm:p-6 lg:p-8 space-y-6">
        {/* Template Status & Metadata Strip */}
        <section className="bg-white border border-slate-200 rounded-xl p-5 shadow-xs transition-all">
          <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
            <div className="flex items-start gap-4">
              <div className="p-3 bg-blue-50 border border-blue-100 rounded-xl text-blue-600 shrink-0">
                <FileSpreadsheet className="w-6 h-6" />
              </div>
              <div className="space-y-1">
                <div className="flex items-center gap-2 flex-wrap">
                  <h2 className="text-base font-bold text-slate-900">
                    {templateMeta?.fileName || '載入中...'}
                  </h2>
                  <span className="text-xs font-semibold px-2 py-0.5 rounded bg-blue-50 text-blue-700 border border-blue-200">
                    {activePreset === 'quote' ? '官方商務報價單範本' : '官方商業請款單範本'}
                  </span>
                </div>
                {/* Metadata with unboxed text and dot separators */}
                <div className="flex items-center gap-2 text-xs text-slate-500 flex-wrap">
                  <span>檔案大小: <span className="font-mono tabular-nums">{templateMeta ? `${(templateMeta.fileSize / 1024).toFixed(1)} KB` : '0 KB'}</span></span>
                  <span aria-hidden="true">·</span>
                  <span>工作表: <span className="font-semibold text-slate-700">{templateMeta?.sheetNames.join(', ') || '無'}</span></span>
                  <span aria-hidden="true">·</span>
                  <span className="flex items-center gap-1 text-emerald-600 font-medium">
                    <Calculator className="w-3.5 h-3.5" />
                    保留 <span className="font-mono tabular-nums">{templateMeta?.formulaCount || 0}</span> 個運算公式
                  </span>
                  <span aria-hidden="true">·</span>
                  <span className="flex items-center gap-1 text-indigo-600 font-medium">
                    <ImageIcon className="w-3.5 h-3.5" />
                    保留 <span className="font-mono tabular-nums">{templateMeta?.imageCount || 0}</span> 個標誌圖片
                  </span>
                </div>
              </div>
            </div>

            {/* Quick Action Badges */}
            <div className="flex items-center gap-2 flex-wrap">
              <button
                type="button"
                onClick={handleResetToDefaults}
                className="px-3 py-1.5 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 border border-slate-200 rounded-lg transition-colors flex items-center gap-1.5"
                title="還原所有欄位為預設值"
              >
                <RotateCcw className="w-3.5 h-3.5 text-slate-500" />
                重設回預設值
              </button>
              <button
                type="button"
                onClick={handleClearAll}
                className="px-3 py-1.5 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 border border-slate-200 rounded-lg transition-colors flex items-center gap-1.5"
                title="清空所有欄位數值"
              >
                <Eraser className="w-3.5 h-3.5 text-slate-500" />
                清空所有欄位
              </button>
              <button
                type="button"
                onClick={handleCopyJson}
                className="px-3 py-1.5 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 border border-slate-200 rounded-lg transition-colors flex items-center gap-1.5"
                title="複製目前填寫內容為 JSON"
              >
                {copiedJson ? (
                  <>
                    <Check className="w-3.5 h-3.5 text-emerald-600" />
                    <span className="text-emerald-700">已複製</span>
                  </>
                ) : (
                  <>
                    <Copy className="w-3.5 h-3.5 text-slate-500" />
                    複製資料 (JSON)
                  </>
                )}
              </button>
            </div>
          </div>

          {/* Quick Metrics Bar */}
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mt-4 pt-4 border-t border-slate-100">
            <div className="bg-slate-50/70 rounded-lg p-2.5 border border-slate-100">
              <span className="text-xs text-slate-500 block">解析出佔位符</span>
              <span className="text-lg font-bold text-slate-800 font-mono tabular-nums">
                {stats.total} <span className="text-xs font-normal text-slate-500">個變數</span>
              </span>
            </div>
            <div className="bg-slate-50/70 rounded-lg p-2.5 border border-slate-100">
              <span className="text-xs text-slate-500 block">已修改欄位</span>
              <span className="text-lg font-bold text-blue-600 font-mono tabular-nums">
                {stats.modified} <span className="text-xs font-normal text-slate-500">個項目</span>
              </span>
            </div>
            <div className="bg-slate-50/70 rounded-lg p-2.5 border border-slate-100">
              <span className="text-xs text-slate-500 block">保留動態公式</span>
              <span className="text-lg font-bold text-emerald-600 font-mono tabular-nums">
                {templateMeta?.formulaCount || 0} <span className="text-xs font-normal text-slate-500">個式子</span>
              </span>
            </div>
            <div className="bg-slate-50/70 rounded-lg p-2.5 border border-slate-100">
              <span className="text-xs text-slate-500 block">未填寫/留空</span>
              <span className="text-lg font-bold text-amber-600 font-mono tabular-nums">
                {stats.empty} <span className="text-xs font-normal text-slate-500">個項目</span>
              </span>
            </div>
          </div>

          {/* Custom Logo Detection & Embedding Options */}
          {hasCustomLogo && (
            <div className="mt-3.5 pt-3.5 border-t border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs bg-slate-50/70 p-3 rounded-lg border border-slate-200/60">
              <div className="flex items-center gap-2.5 text-slate-700">
                <img
                  src={`/logo.png?t=${Date.now()}`}
                  alt="Custom Logo"
                  className="w-7 h-7 object-contain rounded bg-white border border-slate-200 p-0.5 shadow-2xs"
                />
                <div>
                  <div className="flex items-center gap-1.5">
                    <span className="font-bold text-slate-900">
                      已載入專屬商標 (public/logo.png)
                    </span>
                    <span className="text-[10px] bg-emerald-100 text-emerald-800 font-semibold px-1.5 py-0.2 rounded">
                      有效
                    </span>
                  </div>
                  <p className="text-[11px] text-slate-500">
                    支援商務報價單與請款單中的高畫質商標渲染與原樣保留
                  </p>
                </div>
              </div>
              <label className="flex items-center gap-2 cursor-pointer font-medium text-slate-700 select-none shrink-0 bg-white px-3 py-1.5 rounded-md border border-slate-200 shadow-2xs hover:bg-slate-50 transition-colors">
                <input
                  type="checkbox"
                  checked={embedLogoImage}
                  onChange={(e) => setEmbedLogoImage(e.target.checked)}
                  className="rounded border-slate-300 text-blue-600 focus:ring-blue-500 h-3.5 w-3.5"
                />
                <span className="text-xs">匯出時將 logo.png 嵌入至報價單頂部</span>
              </label>
            </div>
          )}
        </section>

        {/* 2 Built-in Preset Templates Switcher Card */}
        <section className="bg-white border border-slate-200 rounded-xl p-3.5 shadow-xs flex flex-col sm:flex-row items-center justify-between gap-3">
          <div className="flex items-center gap-2.5 flex-wrap">
            <span className="text-xs font-bold text-slate-800">選擇 Excel 範本：</span>
            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={() => loadPresetTemplate('quote')}
                className={`px-3.5 py-2 rounded-lg text-xs font-bold transition-all flex items-center gap-2 cursor-pointer ${
                  activePreset === 'quote'
                    ? 'bg-blue-600 text-white shadow-xs ring-2 ring-blue-600/20'
                    : 'bg-slate-100 hover:bg-slate-200 text-slate-700'
                }`}
              >
                <FileText className="w-4 h-4" />
                📄 商務報價單範本
              </button>
              <button
                type="button"
                onClick={() => loadPresetTemplate('invoice')}
                className={`px-3.5 py-2 rounded-lg text-xs font-bold transition-all flex items-center gap-2 cursor-pointer ${
                  activePreset === 'invoice'
                    ? 'bg-blue-600 text-white shadow-xs ring-2 ring-blue-600/20'
                    : 'bg-slate-100 hover:bg-slate-200 text-slate-700'
                }`}
              >
                <FileSpreadsheet className="w-4 h-4" />
                📑 商業請款單範本
              </button>
            </div>
          </div>

          <div className="flex items-center gap-3 flex-wrap">
            <div className="flex items-center gap-1.5 text-xs text-slate-600 bg-slate-50 px-2.5 py-1.5 rounded-lg border border-slate-200">
              <span className="font-semibold text-slate-700">表格排版：</span>
              <label className="flex items-center gap-1 cursor-pointer">
                <input
                  type="radio"
                  name="tableExportMode"
                  value="auto-shrink"
                  checked={tableExportMode === 'auto-shrink'}
                  onChange={() => setTableExportMode('auto-shrink')}
                  className="text-blue-600 focus:ring-blue-500 h-3 w-3"
                />
                <span className="text-xs">依填寫項目動態調整 (緊湊)</span>
              </label>
              <span className="text-slate-300">|</span>
              <label className="flex items-center gap-1 cursor-pointer">
                <input
                  type="radio"
                  name="tableExportMode"
                  value="keep-template-rows"
                  checked={tableExportMode === 'keep-template-rows'}
                  onChange={() => setTableExportMode('keep-template-rows')}
                  className="text-blue-600 focus:ring-blue-500 h-3 w-3"
                />
                <span className="text-xs">固定 5 列範本原貌 (留空未用列)</span>
              </label>
            </div>
            <div className="text-xs text-slate-500 font-medium flex items-center gap-1.5">
              <Sparkles className="w-3.5 h-3.5 text-blue-600" />
              <span>點選立即切換範本</span>
            </div>
          </div>
        </section>

        {/* Control Toolbar: Search, Category Filter, and View Switch */}
        <section className="flex flex-col md:flex-row md:items-center justify-between gap-3">
          {/* Search Input */}
          <div className="relative flex-1 max-w-md">
            <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none" />
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder="搜尋變數名稱、填寫數值或儲存格位址 (例如 B5)..."
              className="w-full pl-9 pr-8 py-2 text-xs bg-white border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all text-slate-800 placeholder:text-slate-400"
            />
            {searchQuery && (
              <button
                type="button"
                onClick={() => setSearchQuery('')}
                className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
              >
                <X className="w-3.5 h-3.5" />
              </button>
            )}
          </div>

          <div className="flex items-center gap-2 flex-wrap justify-between md:justify-end">
            {/* 3 Zones Segmented Control */}
            <div className="flex items-center gap-1 p-1 bg-slate-100 rounded-lg border border-slate-200/60 overflow-x-auto max-w-full">
              <button
                type="button"
                onClick={() => setSelectedZone('all')}
                className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-colors whitespace-nowrap cursor-pointer ${
                  selectedZone === 'all'
                    ? 'bg-white text-slate-900 shadow-xs font-bold'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                📌 全部 3 區（依序）
              </button>
              <button
                type="button"
                onClick={() => setSelectedZone('header')}
                className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-colors whitespace-nowrap flex items-center gap-1.5 cursor-pointer ${
                  selectedZone === 'header'
                    ? 'bg-white text-blue-700 shadow-xs font-bold'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <span className="w-4 h-4 rounded-full bg-blue-100 text-blue-700 text-[10px] font-bold flex items-center justify-center">1</span>
                Header 區
              </button>
              <button
                type="button"
                onClick={() => setSelectedZone('items')}
                className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-colors whitespace-nowrap flex items-center gap-1.5 cursor-pointer ${
                  selectedZone === 'items'
                    ? 'bg-white text-indigo-700 shadow-xs font-bold'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <span className="w-4 h-4 rounded-full bg-indigo-100 text-indigo-700 text-[10px] font-bold flex items-center justify-center">2</span>
                品項區
              </button>
              <button
                type="button"
                onClick={() => setSelectedZone('footer')}
                className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-colors whitespace-nowrap flex items-center gap-1.5 cursor-pointer ${
                  selectedZone === 'footer'
                    ? 'bg-white text-slate-900 shadow-xs font-bold'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <span className="w-4 h-4 rounded-full bg-slate-200 text-slate-700 text-[10px] font-bold flex items-center justify-center">3</span>
                Footer 區
              </button>
            </div>

            {/* View Mode Toggle */}
            <div className="flex items-center gap-1 p-1 bg-slate-100 rounded-lg border border-slate-200/60 shrink-0">
              <button
                type="button"
                onClick={() => setViewMode('form')}
                className={`px-2.5 py-1 text-xs font-medium rounded-md transition-colors flex items-center gap-1.5 ${
                  viewMode === 'form'
                    ? 'bg-white text-blue-700 shadow-xs font-semibold'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
                title="綜合表單輸入模式"
              >
                <Layers className="w-3.5 h-3.5" />
                表單輸入
              </button>
              <button
                type="button"
                onClick={() => setViewMode('items')}
                className={`px-2.5 py-1 text-xs font-medium rounded-md transition-colors flex items-center gap-1.5 ${
                  viewMode === 'items'
                    ? 'bg-white text-blue-700 shadow-xs font-semibold'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
                title="品項明細表格模式（一列代表一個品項）"
              >
                <ListOrdered className="w-3.5 h-3.5" />
                品項表格
                {quotationItems.length > 0 && (
                  <span className="text-[10px] px-1.5 py-0.2 bg-blue-100 text-blue-700 rounded-full font-bold">
                    {quotationItems.length}
                  </span>
                )}
              </button>
              <button
                type="button"
                onClick={() => setViewMode('table')}
                className={`px-2.5 py-1 text-xs font-medium rounded-md transition-colors flex items-center gap-1.5 ${
                  viewMode === 'table'
                    ? 'bg-white text-blue-700 shadow-xs font-semibold'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
                title="儲存格位址對照模式"
              >
                <Table className="w-3.5 h-3.5" />
                儲存格對照表
              </button>
            </div>
          </div>
        </section>

        {/* Content Area: Form View, Items Table View, or Table Mapping View */}
        {isLoading ? (
          <div className="bg-white border border-slate-200 rounded-xl p-16 text-center shadow-xs">
            <RefreshCw className="w-8 h-8 text-blue-600 animate-spin mx-auto mb-3" />
            <p className="text-sm font-semibold text-slate-800">正在分析 Excel 活頁簿結構...</p>
            <p className="text-xs text-slate-500 mt-1">
              掃描儲存格樣式、公式、繪圖與佔位符語法中
            </p>
          </div>
        ) : filteredVariables.length === 0 ? (
          <div className="bg-white border border-slate-200 rounded-xl p-12 text-center shadow-xs">
            <AlertCircle className="w-8 h-8 text-slate-400 mx-auto mb-3" />
            <h3 className="text-sm font-semibold text-slate-800">未找到符合條件的佔位符欄位</h3>
            <p className="text-xs text-slate-500 mt-1">
              {searchQuery ? `找不到與「${searchQuery}」相符的欄位，請嘗試不同關鍵字。` : '此 Excel 檔案未包含任何 {{變數名稱}} 標記。'}
            </p>
            {searchQuery && (
              <button
                type="button"
                onClick={() => setSearchQuery('')}
                className="mt-3 text-xs text-blue-600 font-medium hover:underline"
              >
                清除搜尋條件
              </button>
            )}
          </div>
        ) : viewMode === 'items' ? (
          /* Dedicated Line Items Table Mode */
          <div className="space-y-4">
            {/* Table Component */}
            <div className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs">
              {/* Table Header */}
              <div className="p-4 bg-gradient-to-r from-slate-50 to-blue-50/50 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3">
                <div className="flex items-center gap-2.5">
                  <div className="p-2 bg-blue-600 text-white rounded-lg shadow-2xs">
                    <ListOrdered className="w-4 h-4" />
                  </div>
                  <div>
                    <h3 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                      品項明細清單 (表格模式)
                      <span className="text-[11px] font-semibold text-blue-700 bg-blue-100/70 border border-blue-200 px-2 py-0.5 rounded-full">
                        共 {quotationItems.length} 個品項
                      </span>
                    </h3>
                    <p className="text-xs text-slate-500">
                      一列代表一個品項，數量預設為 1，支援即時小計與動態擴充行數
                    </p>
                  </div>
                </div>

                <button
                  type="button"
                  onClick={handleAddItem}
                  className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg bg-blue-600 hover:bg-blue-700 text-white text-xs font-semibold shadow-xs hover:shadow transition-all active:scale-95 cursor-pointer"
                >
                  <Plus className="w-3.5 h-3.5" />
                  新增品項
                </button>
              </div>

              {/* Table Content */}
              <div className="overflow-x-auto">
                <table className="w-full text-left text-xs border-collapse">
                  <thead>
                    <tr className="bg-slate-100/90 border-b border-slate-200 text-slate-700 font-semibold">
                      <th className="py-2.5 px-3 w-12 text-center">#</th>
                      <th className="py-2.5 px-3 min-w-[180px]">項目名稱</th>
                      <th className="py-2.5 px-3 min-w-[220px]">規格說明與功能簡述</th>
                      <th className="py-2.5 px-3 w-32 text-center">數量 (預設為一個)</th>
                      <th className="py-2.5 px-3 w-28 text-center">單位</th>
                      <th className="py-2.5 px-3 w-32 text-right">單價 (NTD)</th>
                      <th className="py-2.5 px-3 w-32 text-right">小計金額 (NTD)</th>
                      <th className="py-2.5 px-3 w-16 text-center">操作</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {quotationItems.map((item, idx) => {
                      const qtyNum = parseFloat(String(values[item.qtyKey] !== undefined ? values[item.qtyKey] : item.qty)) || 0;
                      const priceNum = parseFloat(String(values[item.priceKey] !== undefined ? values[item.priceKey] : item.price)) || 0;
                      const lineSubtotal = Math.round(qtyNum * priceNum);

                      return (
                        <tr key={item.id} className="hover:bg-blue-50/20 transition-colors">
                          <td className="py-2.5 px-3 text-center font-mono font-medium text-slate-400">
                            {idx + 1}
                          </td>
                          <td className="py-2 px-3">
                            <input
                              type="text"
                              value={values[item.nameKey] !== undefined ? values[item.nameKey] : item.name}
                              onChange={(e) =>
                                setValues((prev) => ({
                                  ...prev,
                                  [item.nameKey]: e.target.value,
                                }))
                              }
                              placeholder="輸入品項名稱"
                              className="w-full text-xs px-2.5 py-1.5 bg-slate-50/70 border border-slate-200 rounded-md focus:bg-white focus:outline-none focus:ring-1 focus:ring-blue-500 focus:border-blue-500 font-medium text-slate-800"
                            />
                          </td>
                          <td className="py-2 px-3">
                            <input
                              type="text"
                              value={values[item.descKey] !== undefined ? values[item.descKey] : item.desc}
                              onChange={(e) =>
                                setValues((prev) => ({
                                  ...prev,
                                  [item.descKey]: e.target.value,
                                }))
                              }
                              placeholder="輸入功能與規格簡述"
                              className="w-full text-xs px-2.5 py-1.5 bg-slate-50/70 border border-slate-200 rounded-md focus:bg-white focus:outline-none focus:ring-1 focus:ring-blue-500 focus:border-blue-500 text-slate-700"
                            />
                          </td>
                          <td className="py-2 px-3">
                            <div className="flex items-center justify-center gap-1">
                              <button
                                type="button"
                                onClick={() => {
                                  const cur = parseFloat(String(values[item.qtyKey] !== undefined ? values[item.qtyKey] : item.qty)) || 1;
                                  if (cur > 1) {
                                    setValues((prev) => ({ ...prev, [item.qtyKey]: String(cur - 1) }));
                                  }
                                }}
                                className="w-5 h-6 flex items-center justify-center rounded bg-slate-100 hover:bg-slate-200 text-slate-600 font-bold border border-slate-200 text-xs cursor-pointer"
                                title="減少數量"
                              >
                                -
                              </button>
                              <input
                                type="number"
                                min="1"
                                step="1"
                                value={values[item.qtyKey] !== undefined ? values[item.qtyKey] : (item.qty || '1')}
                                onChange={(e) =>
                                  setValues((prev) => ({
                                    ...prev,
                                    [item.qtyKey]: e.target.value,
                                  }))
                                }
                                className="w-14 text-xs text-center px-1 py-1.5 bg-slate-50/70 border border-slate-200 rounded-md focus:bg-white focus:outline-none focus:ring-1 focus:ring-blue-500 focus:border-blue-500 font-mono font-semibold text-slate-800"
                              />
                              <button
                                type="button"
                                onClick={() => {
                                  const cur = parseFloat(String(values[item.qtyKey] !== undefined ? values[item.qtyKey] : item.qty)) || 1;
                                  setValues((prev) => ({ ...prev, [item.qtyKey]: String(cur + 1) }));
                                }}
                                className="w-5 h-6 flex items-center justify-center rounded bg-slate-100 hover:bg-slate-200 text-slate-600 font-bold border border-slate-200 text-xs cursor-pointer"
                                title="增加數量"
                              >
                                +
                              </button>
                            </div>
                          </td>
                          <td className="py-2 px-3">
                            <select
                              value={values[item.unitKey] !== undefined ? values[item.unitKey] : item.unit}
                              onChange={(e) =>
                                setValues((prev) => ({
                                  ...prev,
                                  [item.unitKey]: e.target.value,
                                }))
                              }
                              className="w-full text-xs px-2 py-1.5 bg-slate-50/70 border border-slate-200 rounded-md focus:bg-white focus:outline-none focus:ring-1 focus:ring-blue-500 focus:border-blue-500 font-medium text-slate-800 cursor-pointer"
                            >
                              {['專案', '式', '套件', '套', '期', '場次', '個月', '小時', '人天', '個', '組', '台', '份'].map((u) => (
                                <option key={u} value={u}>{u}</option>
                              ))}
                            </select>
                          </td>
                          <td className="py-2 px-3 text-right">
                            <input
                              type="number"
                              min="0"
                              step="100"
                              value={values[item.priceKey] !== undefined ? values[item.priceKey] : item.price}
                              onChange={(e) =>
                                setValues((prev) => ({
                                  ...prev,
                                  [item.priceKey]: e.target.value,
                                }))
                              }
                              className="w-full text-xs text-right px-2.5 py-1.5 bg-slate-50/70 border border-slate-200 rounded-md focus:bg-white focus:outline-none focus:ring-1 focus:ring-blue-500 focus:border-blue-500 font-mono font-medium text-slate-800"
                            />
                          </td>
                          <td className="py-2 px-3 text-right font-mono font-bold text-slate-800">
                            NT$ {lineSubtotal.toLocaleString()}
                          </td>
                          <td className="py-2 px-3 text-center">
                            <button
                              type="button"
                              onClick={() => handleDeleteItem(item)}
                              className="p-1.5 text-slate-400 hover:text-red-600 hover:bg-red-50 rounded transition-colors cursor-pointer"
                              title="刪除此品項"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>

              {/* Financial Summary Footer */}
              <div className="p-4 bg-slate-50 border-t border-slate-200 flex flex-wrap items-center justify-between gap-4">
                <button
                  type="button"
                  onClick={handleAddItem}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-dashed border-blue-400 bg-blue-50/60 hover:bg-blue-100 text-blue-700 text-xs font-semibold transition-colors cursor-pointer"
                >
                  <Plus className="w-3.5 h-3.5" />
                  新增一列品項 (一列代表一個品項)
                </button>

                <div className="flex items-center gap-6 text-xs ml-auto">
                  <div>
                    <span className="text-slate-500 block text-[11px]">銷售額合計 (未稅)</span>
                    <span className="font-mono font-bold text-slate-800 text-sm">
                      NT$ {totalSubtotal.toLocaleString()}
                    </span>
                  </div>
                  <div>
                    <span className="text-slate-500 block text-[11px]">加值型營業稅 (VAT 5%)</span>
                    <span className="font-mono font-semibold text-slate-700 text-sm">
                      NT$ {totalVat.toLocaleString()}
                    </span>
                  </div>
                  <div className="pl-4 border-l border-slate-200">
                    <span className="text-blue-600 block text-[11px] font-bold">總計金額 (含稅)</span>
                    <span className="font-mono font-extrabold text-blue-700 text-base">
                      NT$ {grandTotal.toLocaleString()}
                    </span>
                  </div>
                </div>
              </div>
            </div>
          </div>
        ) : viewMode === 'form' ? (
          /* Form View: Strict 3-Zone Structure (1. Header -> 2. 品項區 -> 3. Footer區) */
          <div className="space-y-8">
            {/* ================= 1. Header 區 ================= */}
            {(selectedZone === 'all' || selectedZone === 'header') && filteredHeaderVars.length > 0 && (
              <section className="bg-white border border-slate-200 rounded-2xl p-5 shadow-xs transition-all space-y-4">
                {/* Header Zone Banner */}
                <div className="flex items-center justify-between border-b border-slate-100 pb-3.5 flex-wrap gap-2">
                  <div className="flex items-center gap-3">
                    <div className="flex items-center justify-center w-8 h-8 rounded-xl bg-blue-600 text-white font-extrabold text-sm shadow-xs">
                      1
                    </div>
                    <div>
                      <h3 className="text-sm font-bold text-slate-900 flex items-center gap-2">
                        Header 區
                        <span className="text-xs font-normal text-slate-500">（單據表頭與基本資訊）</span>
                      </h3>
                      <p className="text-[11px] text-slate-500">
                        包含單號、日期、有效期限、專案名稱、發行公司與客戶買受機構基本資料
                      </p>
                    </div>
                  </div>
                  <span className="text-xs font-semibold px-2.5 py-1 rounded-md bg-blue-50 text-blue-700 border border-blue-200/60 font-mono">
                    共 {filteredHeaderVars.length} 個欄位
                  </span>
                </div>

                {/* Header Fields Grid */}
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  {filteredHeaderVars.map((v) => renderFieldCard(v))}
                </div>
              </section>
            )}

            {/* ================= 2. 品項區 ================= */}
            {(selectedZone === 'all' || selectedZone === 'items') && (
              <section className="space-y-4">
                {/* Items Zone Banner */}
                <div className="bg-white border border-slate-200 rounded-xl px-4 py-3 shadow-xs flex items-center justify-between flex-wrap gap-2">
                  <div className="flex items-center gap-3">
                    <div className="flex items-center justify-center w-8 h-8 rounded-xl bg-indigo-600 text-white font-extrabold text-sm shadow-xs">
                      2
                    </div>
                    <div>
                      <h3 className="text-sm font-bold text-slate-900 flex items-center gap-2">
                        品項區
                        <span className="text-xs font-normal text-slate-500">（品項明細表格）</span>
                      </h3>
                      <p className="text-[11px] text-slate-500">
                        一列代表一個品項，數量預設為 1，支援即時小計、千分位運算與動態新增列
                      </p>
                    </div>
                  </div>
                  <span className="text-xs font-semibold px-2.5 py-1 rounded-md bg-indigo-50 text-indigo-700 border border-indigo-200/60 font-mono">
                    共 {quotationItems.length} 個品項
                  </span>
                </div>

                {/* Line Items Table Component */}
                <div className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs">
                  {/* Table Header */}
                  <div className="p-4 bg-gradient-to-r from-slate-50 to-indigo-50/40 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3">
                    <div className="flex items-center gap-2.5">
                      <div className="p-2 bg-indigo-600 text-white rounded-lg shadow-2xs">
                        <ListOrdered className="w-4 h-4" />
                      </div>
                      <div>
                        <h4 className="text-sm font-bold text-slate-800">
                          品項清單明細
                        </h4>
                        <p className="text-xs text-slate-500">
                          修改數量或單價將即時自動連動計算小計與含稅總金額
                        </p>
                      </div>
                    </div>

                    <button
                      type="button"
                      onClick={handleAddItem}
                      className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold shadow-xs hover:shadow transition-all active:scale-95 cursor-pointer"
                    >
                      <Plus className="w-3.5 h-3.5" />
                      新增品項
                    </button>
                  </div>

                  {/* Table Grid */}
                  <div className="overflow-x-auto">
                    <table className="w-full text-left text-xs border-collapse">
                      <thead>
                        <tr className="bg-slate-100/90 border-b border-slate-200 text-slate-700 font-semibold">
                          <th className="py-2.5 px-3 w-12 text-center">#</th>
                          <th className="py-2.5 px-3 min-w-[180px]">項目名稱</th>
                          <th className="py-2.5 px-3 min-w-[220px]">規格說明與功能簡述</th>
                          <th className="py-2.5 px-3 w-32 text-center">數量 (預設為一個)</th>
                          <th className="py-2.5 px-3 w-28 text-center">單位</th>
                          <th className="py-2.5 px-3 w-32 text-right">單價 (NTD)</th>
                          <th className="py-2.5 px-3 w-32 text-right">小計金額 (NTD)</th>
                          <th className="py-2.5 px-3 w-16 text-center">操作</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {quotationItems.map((item, idx) => {
                          const qtyNum = parseFloat(String(values[item.qtyKey] !== undefined ? values[item.qtyKey] : item.qty)) || 0;
                          const priceNum = parseFloat(String(values[item.priceKey] !== undefined ? values[item.priceKey] : item.price)) || 0;
                          const lineSubtotal = Math.round(qtyNum * priceNum);

                          return (
                            <tr key={item.id} className="hover:bg-indigo-50/20 transition-colors">
                              <td className="py-2.5 px-3 text-center font-mono font-medium text-slate-400">
                                {idx + 1}
                              </td>
                              <td className="py-2 px-3">
                                <input
                                  type="text"
                                  value={values[item.nameKey] !== undefined ? values[item.nameKey] : item.name}
                                  onChange={(e) =>
                                    setValues((prev) => ({
                                      ...prev,
                                      [item.nameKey]: e.target.value,
                                    }))
                                  }
                                  placeholder="輸入品項名稱"
                                  className="w-full text-xs px-2.5 py-1.5 bg-slate-50/70 border border-slate-200 rounded-md focus:bg-white focus:outline-none focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 font-medium text-slate-800"
                                />
                              </td>
                              <td className="py-2 px-3">
                                <input
                                  type="text"
                                  value={values[item.descKey] !== undefined ? values[item.descKey] : item.desc}
                                  onChange={(e) =>
                                    setValues((prev) => ({
                                      ...prev,
                                      [item.descKey]: e.target.value,
                                    }))
                                  }
                                  placeholder="輸入功能與規格簡述"
                                  className="w-full text-xs px-2.5 py-1.5 bg-slate-50/70 border border-slate-200 rounded-md focus:bg-white focus:outline-none focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 text-slate-700"
                                />
                              </td>
                              <td className="py-2 px-3">
                                <div className="flex items-center justify-center gap-1">
                                  <button
                                    type="button"
                                    onClick={() => {
                                      const cur = parseFloat(String(values[item.qtyKey] !== undefined ? values[item.qtyKey] : item.qty)) || 1;
                                      if (cur > 1) {
                                        setValues((prev) => ({ ...prev, [item.qtyKey]: String(cur - 1) }));
                                      }
                                    }}
                                    className="w-5 h-6 flex items-center justify-center rounded bg-slate-100 hover:bg-slate-200 text-slate-600 font-bold border border-slate-200 text-xs cursor-pointer"
                                    title="減少數量"
                                  >
                                    -
                                  </button>
                                  <input
                                    type="number"
                                    min="1"
                                    step="1"
                                    value={values[item.qtyKey] !== undefined ? values[item.qtyKey] : (item.qty || '1')}
                                    onChange={(e) =>
                                      setValues((prev) => ({
                                        ...prev,
                                        [item.qtyKey]: e.target.value,
                                      }))
                                    }
                                    className="w-14 text-xs text-center px-1 py-1.5 bg-slate-50/70 border border-slate-200 rounded-md focus:bg-white focus:outline-none focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 font-mono font-semibold text-slate-800"
                                  />
                                  <button
                                    type="button"
                                    onClick={() => {
                                      const cur = parseFloat(String(values[item.qtyKey] !== undefined ? values[item.qtyKey] : item.qty)) || 1;
                                      setValues((prev) => ({ ...prev, [item.qtyKey]: String(cur + 1) }));
                                    }}
                                    className="w-5 h-6 flex items-center justify-center rounded bg-slate-100 hover:bg-slate-200 text-slate-600 font-bold border border-slate-200 text-xs cursor-pointer"
                                    title="增加數量"
                                  >
                                    +
                                  </button>
                                </div>
                              </td>
                              <td className="py-2 px-3">
                                <select
                                  value={values[item.unitKey] !== undefined ? values[item.unitKey] : item.unit}
                                  onChange={(e) =>
                                    setValues((prev) => ({
                                      ...prev,
                                      [item.unitKey]: e.target.value,
                                    }))
                                  }
                                  className="w-full text-xs px-2 py-1.5 bg-slate-50/70 border border-slate-200 rounded-md focus:bg-white focus:outline-none focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 font-medium text-slate-800 cursor-pointer"
                                >
                                  {['專案', '式', '套件', '套', '期', '場次', '個月', '小時', '人天', '個', '組', '台', '份'].map((u) => (
                                    <option key={u} value={u}>{u}</option>
                                  ))}
                                </select>
                              </td>
                              <td className="py-2 px-3 text-right">
                                <input
                                  type="number"
                                  min="0"
                                  step="100"
                                  value={values[item.priceKey] !== undefined ? values[item.priceKey] : item.price}
                                  onChange={(e) =>
                                    setValues((prev) => ({
                                      ...prev,
                                      [item.priceKey]: e.target.value,
                                    }))
                                  }
                                  className="w-full text-xs text-right px-2.5 py-1.5 bg-slate-50/70 border border-slate-200 rounded-md focus:bg-white focus:outline-none focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 font-mono font-medium text-slate-800"
                                />
                              </td>
                              <td className="py-2 px-3 text-right font-mono font-bold text-slate-800">
                                NT$ {lineSubtotal.toLocaleString()}
                              </td>
                              <td className="py-2 px-3 text-center">
                                <button
                                  type="button"
                                  onClick={() => handleDeleteItem(item)}
                                  className="p-1.5 text-slate-400 hover:text-red-600 hover:bg-red-50 rounded transition-colors cursor-pointer"
                                  title="刪除此品項"
                                >
                                  <Trash2 className="w-3.5 h-3.5" />
                                </button>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>

                  {/* Financial Summary Footer */}
                  <div className="p-4 bg-slate-50 border-t border-slate-200 flex flex-wrap items-center justify-between gap-4">
                    <button
                      type="button"
                      onClick={handleAddItem}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-dashed border-indigo-400 bg-indigo-50/60 hover:bg-indigo-100 text-indigo-700 text-xs font-semibold transition-colors cursor-pointer"
                    >
                      <Plus className="w-3.5 h-3.5" />
                      新增一列品項 (一列代表一個品項)
                    </button>

                    <div className="flex items-center gap-6 text-xs ml-auto">
                      <div>
                        <span className="text-slate-500 block text-[11px]">銷售額合計 (未稅)</span>
                        <span className="font-mono font-bold text-slate-800 text-sm">
                          NT$ {totalSubtotal.toLocaleString()}
                        </span>
                      </div>
                      <div>
                        <span className="text-slate-500 block text-[11px]">加值型營業稅 (VAT 5%)</span>
                        <span className="font-mono font-semibold text-slate-700 text-sm">
                          NT$ {totalVat.toLocaleString()}
                        </span>
                      </div>
                      <div className="pl-4 border-l border-slate-200">
                        <span className="text-indigo-600 block text-[11px] font-bold">總計金額 (含稅)</span>
                        <span className="font-mono font-extrabold text-indigo-700 text-base">
                          NT$ {grandTotal.toLocaleString()}
                        </span>
                      </div>
                    </div>
                  </div>
                </div>
              </section>
            )}

            {/* ================= 3. Footer 區 ================= */}
            {(selectedZone === 'all' || selectedZone === 'footer') && filteredFooterVars.length > 0 && (
              <section className="bg-white border border-slate-200 rounded-2xl p-5 shadow-xs transition-all space-y-4">
                {/* Footer Zone Banner */}
                <div className="flex items-center justify-between border-b border-slate-100 pb-3.5 flex-wrap gap-2">
                  <div className="flex items-center gap-3">
                    <div className="flex items-center justify-center w-8 h-8 rounded-xl bg-slate-800 text-white font-extrabold text-sm shadow-xs">
                      3
                    </div>
                    <div>
                      <h3 className="text-sm font-bold text-slate-900 flex items-center gap-2">
                        Footer 區
                        <span className="text-xs font-normal text-slate-500">（交易條款、金融帳戶與備註）</span>
                      </h3>
                      <p className="text-[11px] text-slate-500">
                        包含付款階段條件、指定金融匯款專戶資訊與交易法律條款備註說明
                      </p>
                    </div>
                  </div>
                  <span className="text-xs font-semibold px-2.5 py-1 rounded-md bg-slate-100 text-slate-700 border border-slate-200 font-mono">
                    共 {filteredFooterVars.length} 個欄位
                  </span>
                </div>

                {/* Footer Fields Grid */}
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  {filteredFooterVars.map((v) => renderFieldCard(v))}
                </div>
              </section>
            )}
          </div>
        ) : (
          /* Table Mapping View */
          <div className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs">
            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs">
                <thead>
                  <tr className="bg-slate-100/80 border-b border-slate-200 text-slate-600 font-semibold">
                    <th className="py-2.5 px-4 w-12 text-center">#</th>
                    <th className="py-2.5 px-4">變數名稱 (Key)</th>
                    <th className="py-2.5 px-4">工作表 / 儲存格</th>
                    <th className="py-2.5 px-4">範本預設值</th>
                    <th className="py-2.5 px-4">當前填寫值</th>
                    <th className="py-2.5 px-4 w-28 text-center">輸入控制項</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 font-sans">
                  {filteredVariables.map((v, idx) => {
                    const isModified = values[v.key] !== v.defaultValue;
                    return (
                      <tr key={v.key} className="hover:bg-slate-50/80 transition-colors">
                        <td className="py-2 px-4 text-center font-mono text-slate-400 tabular-nums">
                          {idx + 1}
                        </td>
                        <td className="py-2 px-4 font-semibold text-slate-800">
                          <div>{formatKeyToChinese(v.key)}</div>
                          {formatKeyToChinese(v.key) !== v.key && (
                            <div className="text-[10px] font-mono text-slate-400">{v.key}</div>
                          )}
                        </td>
                        <td className="py-2 px-4 font-mono text-slate-500 text-[11px]">
                          {v.occurrences.map((occ) => (
                            <span key={occ.cellAddress} className="inline-block bg-slate-100 border border-slate-200/70 rounded px-1.5 py-0.5 mr-1 mb-1">
                              {occ.sheetName}!{occ.cellAddress}
                            </span>
                          ))}
                        </td>
                        <td className="py-2 px-4 text-slate-500 font-mono truncate max-w-xs">
                          {v.defaultValue || '—'}
                        </td>
                        <td className="py-2 px-4">
                          {v.isDate ? (
                            <input
                              type="date"
                              value={normalizeDateValue(values[v.key] || '')}
                              onChange={(e) =>
                                setValues((prev) => ({
                                  ...prev,
                                  [v.key]: e.target.value,
                                }))
                              }
                              className={`w-full text-xs px-2.5 py-1 rounded border transition-all cursor-pointer font-mono font-medium ${
                                isModified
                                  ? 'bg-blue-50/50 border-blue-300 font-medium text-blue-900'
                                  : 'bg-slate-50 border-slate-200 text-slate-800'
                              }`}
                            />
                          ) : v.dropdownOptions && v.dropdownOptions.length > 0 ? (
                            <div className="flex items-center gap-1.5">
                              <select
                                value={
                                  v.dropdownOptions.includes(values[v.key] ?? '')
                                    ? (values[v.key] ?? '')
                                    : '__custom__'
                                }
                                onChange={(e) => {
                                  if (e.target.value !== '__custom__') {
                                    setValues((prev) => ({
                                      ...prev,
                                      [v.key]: e.target.value,
                                    }));
                                  }
                                }}
                                className={`text-xs px-2 py-1 rounded border transition-all cursor-pointer font-medium ${
                                  isModified
                                    ? 'bg-blue-50/50 border-blue-300 font-medium text-blue-900'
                                    : 'bg-white border-slate-200 text-slate-800'
                                }`}
                              >
                                <option value="" disabled>-- 下拉選單 --</option>
                                {v.dropdownOptions.map((opt) => (
                                  <option key={opt} value={opt}>
                                    {opt}
                                  </option>
                                ))}
                                <option value="__custom__">✏️ 自訂輸入</option>
                              </select>
                              <input
                                type="text"
                                value={values[v.key] || ''}
                                onChange={(e) =>
                                  setValues((prev) => ({
                                    ...prev,
                                    [v.key]: e.target.value,
                                  }))
                                }
                                className={`flex-1 text-xs px-2 py-1 rounded border transition-all ${
                                  isModified
                                    ? 'bg-blue-50/50 border-blue-300 font-medium text-blue-900'
                                    : 'bg-slate-50 border-slate-200 text-slate-800'
                                }`}
                              />
                            </div>
                          ) : (
                            <input
                              type="text"
                              value={values[v.key] || ''}
                              onChange={(e) =>
                                setValues((prev) => ({
                                  ...prev,
                                  [v.key]: e.target.value,
                                }))
                              }
                              className={`w-full text-xs px-2.5 py-1 rounded border transition-all ${
                                isModified
                                  ? 'bg-blue-50/50 border-blue-300 font-medium text-blue-900'
                                  : 'bg-slate-50 border-slate-200 text-slate-800'
                              }`}
                            />
                          )}
                        </td>
                        <td className="py-2 px-4 text-center">
                          <span className="text-[11px] font-mono text-slate-500 bg-slate-100 px-2 py-0.5 rounded border border-slate-200">
                            {v.isDate ? '日曆日期' : v.dropdownOptions ? '下拉選單' : v.isLongText ? '多行文字' : '單行文字'}
                          </span>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* Feature Explanation & Architecture Note */}
        <section className="bg-slate-100/70 border border-slate-200/80 rounded-xl p-5 text-xs text-slate-600 space-y-3">
          <div className="flex items-center gap-2 font-bold text-slate-800 text-sm">
            <Sparkles className="w-4 h-4 text-blue-600" />
            <span>核心架構與功能保障說明</span>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            <div className="bg-white p-3.5 rounded-lg border border-slate-200/60 shadow-xs space-y-1">
              <span className="font-semibold text-slate-800 block">1. 完整保留圖片與版面樣式</span>
              <p className="text-slate-500 leading-relaxed">
                採用 ExcelJS 雙向緩衝區架構，匯出時完全保留原範本內的企業 Logo、頁首頁尾、格線顏色與欄寬列高設定。
              </p>
            </div>
            <div className="bg-white p-3.5 rounded-lg border border-slate-200/60 shadow-xs space-y-1">
              <span className="font-semibold text-slate-800 block">2. 動態公式即時連動</span>
              <p className="text-slate-500 leading-relaxed">
                數字型佔位符在匯出時自動轉為數值型別（Numeric），確保所有 Excel 公式（如 <code className="bg-slate-100 px-1 py-0.5 rounded text-blue-600 font-mono">=D*E</code>、<code className="bg-slate-100 px-1 py-0.5 rounded text-blue-600 font-mono">=SUM(...)</code>）皆可自動計算。
              </p>
            </div>
            <div className="bg-white p-3.5 rounded-lg border border-slate-200/60 shadow-xs space-y-1">
              <span className="font-semibold text-slate-800 block">3. 純前端免伺服器安全機制</span>
              <p className="text-slate-500 leading-relaxed">
                100% 在本機瀏覽器內完成解析與合成，零 API Key 依賴，無任何資料外洩風險，亦可離線運作。
              </p>
            </div>
          </div>
        </section>
      </main>

      {/* Footer */}
      <footer className="bg-white border-t border-slate-200 px-6 py-4 text-center text-xs text-slate-500">
        <p>通用型 Excel 佔位符單據產生器 · Universal Template Engine · 純前端免伺服器運作</p>
      </footer>

      {/* Floating Toast Notification Stack */}
      <div className="fixed bottom-5 right-5 z-50 flex flex-col gap-2 max-w-sm w-full pointer-events-none">
        {toasts.map((t) => (
          <div
            key={t.id}
            className={`pointer-events-auto p-3.5 rounded-xl shadow-lg border text-xs flex items-start gap-3 transition-all animate-in fade-in slide-in-from-bottom-2 ${
              t.type === 'success'
                ? 'bg-white border-emerald-200 text-emerald-950'
                : t.type === 'error'
                ? 'bg-white border-rose-200 text-rose-950'
                : t.type === 'warning'
                ? 'bg-white border-amber-200 text-amber-950'
                : 'bg-white border-blue-200 text-slate-900'
            }`}
          >
            <div className="shrink-0 mt-0.5">
              {t.type === 'success' && <CheckCircle2 className="w-4 h-4 text-emerald-600" />}
              {t.type === 'error' && <AlertCircle className="w-4 h-4 text-rose-600" />}
              {t.type === 'warning' && <AlertCircle className="w-4 h-4 text-amber-600" />}
              {t.type === 'info' && <Info className="w-4 h-4 text-blue-600" />}
            </div>
            <div className="flex-1 space-y-0.5">
              <p className="font-bold text-slate-900">{t.title}</p>
              <p className="text-slate-600 leading-relaxed">{t.message}</p>
            </div>
            <button
              type="button"
              onClick={() => removeToast(t.id)}
              className="text-slate-400 hover:text-slate-600 p-0.5"
            >
              <X className="w-3.5 h-3.5" />
            </button>
          </div>
        ))}
      </div>
    </div>
  );
}
