import React, { useState, useEffect } from 'react';
import {
  TableCellsMerge,
  Plus,
  Wand2,
  Trash2,
  Download,
  CheckCircle2,
  X,
  Keyboard,
  AlertCircle,
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { SheetRow } from '../types';

interface ManualEditorModalProps {
  isOpen: boolean;
  onClose: () => void;
  initialMode: 1 | 2;
  currentSheet1: SheetRow[];
  currentSheet2: SheetRow[];
  onApplySheet1: (data: SheetRow[]) => void;
  onApplySheet2: (data: SheetRow[]) => void;
}

interface EditorRowItem {
  c1: string;
  c2: string;
  c3?: string;
}

export const ManualEditorModal: React.FC<ManualEditorModalProps> = ({
  isOpen,
  onClose,
  initialMode,
  currentSheet1,
  currentSheet2,
  onApplySheet1,
  onApplySheet2,
}) => {
  const [activeMode, setActiveMode] = useState<1 | 2>(initialMode);
  const [editorData, setEditorData] = useState<EditorRowItem[]>([]);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [showClearConfirm, setShowClearConfirm] = useState(false);

  useEffect(() => {
    setActiveMode(initialMode);
    setErrorMessage(null);
    setShowClearConfirm(false);
  }, [initialMode, isOpen]);

  useEffect(() => {
    if (!isOpen) return;

    if (activeMode === 1) {
      if (currentSheet1.length > 0) {
        setEditorData(
          currentSheet1.map((r) => ({
            c1: String(r['رقم الماكينة'] ?? Object.values(r)[0] ?? ''),
            c2: String(r['الفرع / المنطقة'] ?? Object.values(r)[1] ?? ''),
          }))
        );
      } else {
        setEditorData([
          { c1: 'POS-1001', c2: 'فرع القاهرة - المعادي' },
          { c1: 'POS-1002', c2: 'فرع الجيزة - الدقي' },
          { c1: 'POS-1003', c2: 'فرع الإسكندرية - سموحة' },
          { c1: 'POS-1004', c2: 'فرع المنصورة - المشاية' },
          { c1: 'POS-1005', c2: 'فرع طنطا - المحطة' },
          { c1: 'POS-1006', c2: 'فرع أسيوط - الجمهورية' },
        ]);
      }
    } else {
      if (currentSheet2.length > 0) {
        setEditorData(
          currentSheet2.map((r) => ({
            c1: String(r['رقم الماكينة'] ?? Object.values(r)[0] ?? ''),
            c2: String(r['رقم حساب المندوب'] ?? Object.values(r)[1] ?? ''),
            c3: String(r['اسم المندوب'] ?? Object.values(r)[2] ?? ''),
          }))
        );
      } else {
        setEditorData([
          { c1: 'POS-1001', c2: '1234', c3: 'أحمد محمود سالم' },
          { c1: 'POS-1001', c2: '456', c3: 'محمود حسن رضوان' },
          { c1: 'POS-1002', c2: '555', c3: 'خالد عبد الرحمن' },
          { c1: 'POS-1003', c2: '789', c3: 'طارق زياد العتيبي' },
          { c1: 'POS-1003', c2: '890', c3: 'عمر فاروق الشامي' },
          { c1: 'POS-1003', c2: '999', c3: 'إبراهيم حسني مراد' },
          { c1: 'POS-1004', c2: '777', c3: 'سامح عبد الله كمال' },
        ]);
      }
    }
  }, [activeMode, isOpen, currentSheet1, currentSheet2]);

  if (!isOpen) return null;

  const handleCellChange = (idx: number, field: 'c1' | 'c2' | 'c3', val: string) => {
    setErrorMessage(null);
    setEditorData((prev) => {
      const next = [...prev];
      next[idx] = { ...next[idx], [field]: val };
      return next;
    });
  };

  const addRow = () => {
    setErrorMessage(null);
    setEditorData((prev) => [...prev, { c1: '', c2: '', c3: '' }]);
  };

  const removeRow = (idx: number) => {
    setErrorMessage(null);
    setEditorData((prev) => prev.filter((_, i) => i !== idx));
  };

  const confirmClearAll = () => {
    setEditorData([]);
    setShowClearConfirm(false);
    setErrorMessage(null);
  };

  const loadDefaults = () => {
    setErrorMessage(null);
    setShowClearConfirm(false);
    if (activeMode === 1) {
      setEditorData([
        { c1: 'POS-1001', c2: 'فرع القاهرة - المعادي' },
        { c1: 'POS-1002', c2: 'فرع الجيزة - الدقي' },
        { c1: 'POS-1003', c2: 'فرع الإسكندرية - سموحة' },
        { c1: 'POS-1004', c2: 'فرع المنصورة - المشاية' },
        { c1: 'POS-1005', c2: 'فرع طنطا - المحطة' },
        { c1: 'POS-1006', c2: 'فرع أسيوط - الجمهورية' },
      ]);
    } else {
      setEditorData([
        { c1: 'POS-1001', c2: '1234', c3: 'أحمد محمود سالم' },
        { c1: 'POS-1001', c2: '456', c3: 'محمود حسن رضوان' },
        { c1: 'POS-1002', c2: '555', c3: 'خالد عبد الرحمن' },
        { c1: 'POS-1003', c2: '789', c3: 'طارق زياد العتيبي' },
        { c1: 'POS-1003', c2: '890', c3: 'عمر فاروق الشامي' },
        { c1: 'POS-1003', c2: '999', c3: 'إبراهيم حسني مراد' },
        { c1: 'POS-1004', c2: '777', c3: 'سامح عبد الله كمال' },
      ]);
    }
  };

  const downloadAsExcel = () => {
    if (!editorData.length) {
      setErrorMessage('الجدول فارغ حالياً! أضف أسطر أو اضغط على إعادة تعبئة أمثلة تجريبية.');
      return;
    }
    const filename =
      activeMode === 1 ? 'قالب_الماكينات_المعدل_يدوياً' : 'قالب_المناديب_المعدل_يدوياً';
    const aoa: (string | number)[][] = [];

    if (activeMode === 1) {
      aoa.push(['رقم الماكينة', 'الفرع / المنطقة']);
      editorData.forEach((r) => aoa.push([r.c1 || '', r.c2 || '']));
    } else {
      aoa.push(['رقم الماكينة', 'رقم حساب المندوب', 'اسم المندوب']);
      editorData.forEach((r) => aoa.push([r.c1 || '', r.c2 || '', r.c3 || '']));
    }

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    XLSX.utils.book_append_sheet(wb, ws, 'البيانات');
    XLSX.writeFile(wb, `${filename}.xlsx`);
  };

  const applyToSystem = () => {
    const valid = editorData.filter((r) => r.c1 && r.c1.trim() !== '');
    if (!valid.length) {
      setErrorMessage('يرجى التأكد من كتابة رقم الماكينة في سطر واحد على الأقل قبل الاعتماد.');
      return;
    }

    if (activeMode === 1) {
      const converted: SheetRow[] = valid.map((r) => ({
        'رقم الماكينة': r.c1.trim(),
        'الفرع / المنطقة': r.c2?.trim() || '',
      }));
      onApplySheet1(converted);
    } else {
      const converted: SheetRow[] = valid.map((r) => ({
        'رقم الماكينة': r.c1.trim(),
        'رقم حساب المندوب': r.c2?.trim() || '',
        'اسم المندوب': r.c3?.trim() || '',
      }));
      onApplySheet2(converted);
    }
    onClose();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/80 backdrop-blur-md p-3 sm:p-5 transition-all">
      <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-3xl w-full max-w-5xl shadow-2xl overflow-hidden flex flex-col max-h-[92vh]">
        {/* Header */}
        <div className="p-5 border-b border-slate-200 dark:border-slate-800 flex items-center justify-between bg-slate-50/50 dark:bg-slate-800/40">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-2xl bg-indigo-600 text-white flex items-center justify-center text-lg shadow-lg shadow-indigo-500/25">
              <TableCellsMerge className="w-5 h-5" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h3 className="font-black text-slate-900 dark:text-white text-base">
                  محرر القوالب التفاعلي أوفلاين (Online/Offline Table Editor)
                </h3>
                <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-indigo-100 dark:bg-indigo-950 text-indigo-600 dark:text-indigo-400">
                  {activeMode === 1 ? 'قالب الماكينات (Sheet 1)' : 'قالب المناديب (Sheet 2)'}
                </span>
              </div>
              <p className="text-xs text-slate-500 dark:text-slate-400">
                أدخل أو عدّل البيانات مباشرة في الجدول، ثم نزّل شيت إكسل جاهز أو اعتمدها فوراً في المطابقة!
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="w-9 h-9 rounded-xl bg-slate-100 dark:bg-slate-800 hover:bg-rose-50 hover:text-rose-600 dark:hover:bg-rose-950/50 text-slate-500 dark:text-slate-400 flex items-center justify-center transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Mode Switcher & Tools */}
        <div className="px-6 py-3 bg-white dark:bg-slate-900 border-b border-slate-200 dark:border-slate-800 flex flex-wrap items-center justify-between gap-3">
          <div className="flex items-center gap-2">
            <button
              onClick={() => {
                setActiveMode(1);
                setErrorMessage(null);
              }}
              className={`px-3.5 py-1.5 rounded-xl text-xs font-bold transition-all ${
                activeMode === 1
                  ? 'bg-indigo-600 text-white shadow-sm'
                  : 'text-slate-600 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-800'
              }`}
            >
              قالب شيت الماكينات (Sheet 1)
            </button>
            <button
              onClick={() => {
                setActiveMode(2);
                setErrorMessage(null);
              }}
              className={`px-3.5 py-1.5 rounded-xl text-xs font-bold transition-all ${
                activeMode === 2
                  ? 'bg-blue-600 text-white shadow-sm'
                  : 'text-slate-600 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-800'
              }`}
            >
              قالب شيت المناديب والحسابات (Sheet 2)
            </button>
          </div>

          <div className="flex items-center gap-2 flex-wrap">
            <button
              onClick={addRow}
              className="px-3 py-1.5 rounded-xl text-xs font-bold bg-slate-100 hover:bg-slate-200 dark:bg-slate-800 dark:hover:bg-slate-700 text-slate-700 dark:text-slate-300 flex items-center gap-1.5 transition-all"
            >
              <Plus className="w-3.5 h-3.5 text-emerald-500" />
              <span>إضافة سطر جديد</span>
            </button>
            <button
              onClick={loadDefaults}
              className="px-3 py-1.5 rounded-xl text-xs font-bold bg-slate-100 hover:bg-slate-200 dark:bg-slate-800 dark:hover:bg-slate-700 text-amber-600 dark:text-amber-400 flex items-center gap-1.5 transition-all"
            >
              <Wand2 className="w-3.5 h-3.5" />
              <span>إعادة تعبئة أمثلة تجريبية</span>
            </button>

            {showClearConfirm ? (
              <div className="flex items-center gap-1.5 bg-rose-50 dark:bg-rose-950/60 p-1 rounded-xl border border-rose-200 dark:border-rose-900/60">
                <span className="text-[11px] font-bold text-rose-600 dark:text-rose-400 px-1">
                  تأكيد المسح؟
                </span>
                <button
                  onClick={confirmClearAll}
                  className="px-2 py-0.5 bg-rose-600 text-white rounded-lg text-xs font-bold"
                >
                  نعم
                </button>
                <button
                  onClick={() => setShowClearConfirm(false)}
                  className="px-2 py-0.5 bg-slate-200 dark:bg-slate-700 text-slate-700 dark:text-slate-200 rounded-lg text-xs font-bold"
                >
                  إلغاء
                </button>
              </div>
            ) : (
              <button
                onClick={() => setShowClearConfirm(true)}
                className="px-3 py-1.5 rounded-xl text-xs font-bold bg-rose-50 hover:bg-rose-100 dark:bg-rose-950/40 text-rose-600 dark:text-rose-400 flex items-center gap-1.5 transition-all"
              >
                <Trash2 className="w-3.5 h-3.5" />
                <span>تفريغ الجدول</span>
              </button>
            )}
          </div>
        </div>

        {/* Error Notification Banner if any */}
        {errorMessage && (
          <div className="mx-6 mt-3 p-3 rounded-2xl bg-rose-50 dark:bg-rose-950/60 border border-rose-200 dark:border-rose-900 text-rose-700 dark:text-rose-300 text-xs font-bold flex items-center justify-between gap-2 animate-in fade-in">
            <div className="flex items-center gap-2">
              <AlertCircle className="w-4 h-4 text-rose-500 shrink-0" />
              <span>{errorMessage}</span>
            </div>
            <button
              onClick={() => setErrorMessage(null)}
              className="text-rose-500 hover:text-rose-700 font-bold"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        )}

        {/* Table Body Area */}
        <div className="flex-1 overflow-auto p-6 space-y-3 bg-slate-50/50 dark:bg-[#070b13]">
          <div className="overflow-x-auto rounded-2xl border border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900 shadow-sm">
            <table className="w-full text-right border-collapse text-xs">
              <thead className="bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300 font-bold border-b border-slate-200 dark:border-slate-800">
                {activeMode === 1 ? (
                  <tr>
                    <th className="p-3 w-12 text-center">م</th>
                    <th className="p-3 w-1/2">رقم الماكينة (Machine ID) *</th>
                    <th className="p-3 w-1/2">الفرع / المنطقة (اختياري)</th>
                    <th className="p-3 w-20 text-center">إجراء</th>
                  </tr>
                ) : (
                  <tr>
                    <th className="p-3 w-12 text-center">م</th>
                    <th className="p-3 w-1/3">رقم الماكينة *</th>
                    <th className="p-3 w-1/4">رقم حساب المندوب *</th>
                    <th className="p-3 w-1/3">اسم المندوب المسند</th>
                    <th className="p-3 w-20 text-center">إجراء</th>
                  </tr>
                )}
              </thead>
              <tbody className="divide-y divide-slate-100 dark:divide-slate-800">
                {editorData.length === 0 ? (
                  <tr>
                    <td colSpan={5} className="p-8 text-center text-slate-400">
                      الجدول فارغ حالياً. اضغط على "إضافة سطر جديد" أو "إعادة تعبئة أمثلة تجريبية".
                    </td>
                  </tr>
                ) : (
                  editorData.map((row, idx) => (
                    <tr key={idx} className="hover:bg-slate-50 dark:hover:bg-slate-800/40 transition-colors">
                      <td className="p-2 text-center text-slate-400 font-mono">{idx + 1}</td>
                      <td className="p-2">
                        <input
                          type="text"
                          className="w-full px-3 py-1.5 bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-xl text-xs font-mono font-bold text-slate-800 dark:text-slate-100 focus:outline-none focus:border-indigo-500"
                          value={row.c1}
                          onChange={(e) => handleCellChange(idx, 'c1', e.target.value)}
                          placeholder="مثال: POS-1001"
                        />
                      </td>
                      <td className="p-2">
                        <input
                          type="text"
                          className="w-full px-3 py-1.5 bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-xl text-xs font-mono font-black text-indigo-600 dark:text-indigo-400 focus:outline-none focus:border-amber-500"
                          value={row.c2}
                          onChange={(e) => handleCellChange(idx, 'c2', e.target.value)}
                          placeholder={activeMode === 1 ? 'مثال: فرع القاهرة' : 'مثال: 1234'}
                        />
                      </td>
                      {activeMode === 2 && (
                        <td className="p-2">
                          <input
                            type="text"
                            className="w-full px-3 py-1.5 bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-xl text-xs font-semibold text-slate-800 dark:text-slate-100 focus:outline-none focus:border-emerald-500"
                            value={row.c3 || ''}
                            onChange={(e) => handleCellChange(idx, 'c3', e.target.value)}
                            placeholder="اسم المندوب أو اتركه فارغاً"
                          />
                        </td>
                      )}
                      <td className="p-2 text-center">
                        <button
                          onClick={() => removeRow(idx)}
                          className="w-7 h-7 rounded-lg text-rose-500 hover:bg-rose-50 dark:hover:bg-rose-950/40 inline-flex items-center justify-center transition-colors"
                          title="حذف هذا السطر"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
          <div className="flex items-center justify-between text-[11px] text-slate-500 dark:text-slate-400 pt-1 px-1">
            <span>إجمالي الأسطر: {editorData.length}</span>
            <span className="text-indigo-500 dark:text-indigo-400 font-semibold flex items-center gap-1">
              <Keyboard className="w-3.5 h-3.5" />
              <span>يمكنك الكتابة مباشرة داخل الحقول أو لصق أرقام الماكينات والبيانات.</span>
            </span>
          </div>
        </div>

        {/* Footer */}
        <div className="p-4 sm:p-5 border-t border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900 flex flex-wrap items-center justify-between gap-3">
          <button
            onClick={onClose}
            className="px-4 py-2 rounded-xl text-xs font-bold text-slate-500 hover:text-slate-700 dark:text-slate-400 dark:hover:text-slate-200 transition-colors"
          >
            إلغاء وإغلاق
          </button>
          <div className="flex items-center gap-2">
            <button
              onClick={downloadAsExcel}
              className="px-4 py-2 rounded-xl text-xs font-bold bg-slate-100 hover:bg-slate-200 dark:bg-slate-800 dark:hover:bg-slate-700 text-slate-800 dark:text-slate-200 border border-slate-200 dark:border-slate-700 flex items-center gap-1.5 transition-all shadow-sm"
            >
              <Download className="w-4 h-4 text-emerald-500" />
              <span>تنزيل ملف الإكسل (.xlsx)</span>
            </button>
            <button
              onClick={applyToSystem}
              className="px-5 py-2 rounded-xl text-xs font-bold bg-indigo-600 hover:bg-indigo-500 text-white flex items-center gap-2 shadow-lg shadow-indigo-600/25 transition-all"
            >
              <CheckCircle2 className="w-4 h-4 text-amber-300" />
              <span>اعتماد واستخدام البيانات فوراً في المطابقة</span>
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
