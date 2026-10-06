import React, { useState, useEffect, useMemo, useRef } from 'react';
import {
  Layers,
  FileCheck,
  FileSpreadsheet,
  Download,
  Sliders,
  CheckCheck,
  ArrowRight,
  Search,
  Users,
  UserCheck,
  CircleAlert,
  TableProperties,
  ArrowDown,
  UserX,
  UploadCloud,
  ChevronLeft,
  ChevronRight,
  Monitor,
  Zap,
  CheckCircle2,
  X,
  User,
  FileCode,
  Calculator,
} from 'lucide-react';
import { SheetRow, ExpandedRow, MachineSummary, DuplicateAuditItem, IrregularAccountItem } from './types';
import { parseExcelFile, downloadTemplateFile, exportReconciliationToExcel } from './utils/excel';
import { DesktopHeader } from './components/DesktopHeader';
import { DesktopAppCenterModal } from './components/DesktopAppCenterModal';
import { ManualEditorModal } from './components/ManualEditorModal';
import { ResetConfirmModal } from './components/ResetConfirmModal';
import { RepLookupModal } from './components/RepLookupModal';
import { AuditConflictsTab } from './components/AuditConflictsTab';
import { BulkActionBar } from './components/BulkActionBar';
import { ProgressModal } from './components/ProgressModal';
import { OfflineIndicator } from './components/OfflineIndicator';
import { DesktopAppTab } from './components/DesktopAppTab';

export default function App() {
  // Theme state with localStorage persistence
  const [isDarkMode, setIsDarkMode] = useState<boolean>(() => {
    try {
      const saved = localStorage.getItem('theme');
      if (saved) return saved === 'dark';
    } catch (e) {}
    return true; // Default dark
  });

  // Tab state (1 to 5)
  const [currentTab, setCurrentTab] = useState<number>(1);

  // Modals state
  const [isDesktopCenterOpen, setIsDesktopCenterOpen] = useState(false);
  const [isManualEditorOpen, setIsManualEditorOpen] = useState(false);
  const [isResetConfirmOpen, setIsResetConfirmOpen] = useState(false);
  const [isRepLookupOpen, setIsRepLookupOpen] = useState(false);
  const [repLookupQuery, setRepLookupQuery] = useState('');
  const [editorInitialMode, setEditorInitialMode] = useState<1 | 2>(1);

  // Floating Toast Notification
  const [toastMessage, setToastMessage] = useState<string | null>(null);
  const toastTimeoutRef = useRef<NodeJS.Timeout | null>(null);

  const showToast = (msg: string) => {
    if (toastTimeoutRef.current) clearTimeout(toastTimeoutRef.current);
    setToastMessage(msg);
    toastTimeoutRef.current = setTimeout(() => {
      setToastMessage(null);
    }, 3200);
  };

  // Sheets data
  const [sheet1, setSheet1] = useState<SheetRow[]>([]);
  const [sheet2, setSheet2] = useState<SheetRow[]>([]);
  const [sheet1FileName, setSheet1FileName] = useState<string>('');
  const [sheet2FileName, setSheet2FileName] = useState<string>('');

  // Column mapping
  const [colM1, setColM1] = useState<string>('');
  const [colM2, setColM2] = useState<string>('');
  const [colAcc, setColAcc] = useState<string>('');
  const [colRep, setColRep] = useState<string>('');
  const [emptyRepFallback, setEmptyRepFallback] = useState<string>('لا يوجد مندوب');

  // Reconciliation results
  const [machinesResults, setMachinesResults] = useState<MachineSummary[]>([]);
  const [expandedRows, setExpandedRows] = useState<ExpandedRow[]>([]);

  // Filtering & Pagination
  const [currentFilter, setCurrentFilter] = useState<'all' | 'single' | 'multi' | 'none'>('all');
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [currentPage, setCurrentPage] = useState<number>(1);
  const [viewMode, setViewMode] = useState<'cards' | 'table'>('cards'); // Default to vibrant Cards & Boxes view
  const pageSize = 50;
  const searchInputRef = useRef<HTMLInputElement>(null);

  // Bulk actions selection
  const [selectedMachines, setSelectedMachines] = useState<Set<string>>(new Set());

  // Audits
  const [duplicates, setDuplicates] = useState<DuplicateAuditItem[]>([]);
  const [irregulars, setIrregulars] = useState<IrregularAccountItem[]>([]);

  // Progress modal
  const [progressState, setProgressState] = useState({
    isOpen: false,
    title: 'قراءة',
    fileName: '',
    subtitle: 'جاري تحليل خلايا وسجلات ملف الإكسل...',
    percent: 0,
    stepText: 'تحويل ورقة العمل إلى هيكل بيانات JSON...',
    countText: '',
    isComplete: false,
  });

  // Sync theme with HTML class
  useEffect(() => {
    const root = document.documentElement;
    if (isDarkMode) {
      root.classList.add('dark');
    } else {
      root.classList.remove('dark');
    }
  }, [isDarkMode]);

  const handleToggleTheme = () => {
    setIsDarkMode((prev) => {
      const next = !prev;
      try {
        localStorage.setItem('theme', next ? 'dark' : 'light');
      } catch (e) {}
      if (next) {
        document.documentElement.classList.add('dark');
      } else {
        document.documentElement.classList.remove('dark');
      }
      showToast(next ? '🌙 تم تفعيل الوضع الليلي' : '☀️ تم تفعيل الوضع النهاري');
      return next;
    });
  };

  // Global Keyboard Shortcuts
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'd') {
        e.preventDefault();
        setIsDesktopCenterOpen((prev) => !prev);
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'e') {
        e.preventDefault();
        setEditorInitialMode(1);
        setIsManualEditorOpen(true);
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
        e.preventDefault();
        if (expandedRows.length > 0) {
          handleExportExcel();
        }
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'f') {
        e.preventDefault();
        setCurrentTab(4);
        setTimeout(() => searchInputRef.current?.focus(), 100);
      }
      if (e.key === 'Escape') {
        setIsDesktopCenterOpen(false);
        setIsManualEditorOpen(false);
        setIsResetConfirmOpen(false);
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [expandedRows]);

  // Auto-detect columns whenever sheets change
  useEffect(() => {
    if (sheet1.length > 0) {
      const cols1 = Object.keys(sheet1[0]);
      const foundM1 = cols1.find((c) => /ماكينة|جهاز|pos|machine|sn|id/i.test(c)) || cols1[0];
      setColM1(foundM1);
    }

    if (sheet2.length > 0) {
      const cols2 = Object.keys(sheet2[0]);
      const foundM2 = cols2.find((c) => /ماكينة|جهاز|pos|machine|sn/i.test(c)) || cols2[0];
      const foundAcc = cols2.find((c) => /حساب|كود|رقم.*مندوب|account|acc|code/i.test(c)) || '';
      let foundRep = cols2.find((c) => /اسم.*مندوب|اسم.*العميل|اسم.*المستخدم|اسم|rep.*name|agent.*name|name/i.test(c));
      if (!foundRep) {
        foundRep = cols2.find((c) => /(مندوب|agent|rep)/i.test(c) && !/(رقم|حساب|كود|acc|code|id)/i.test(c));
      }
      if (!foundRep) {
        foundRep = cols2.find((c) => /(مندوب|agent|rep)/i.test(c)) || cols2[1] || cols2[0];
      }

      setColM2(foundM2);
      setColAcc(foundAcc);
      setColRep(foundRep);
    }
  }, [sheet1, sheet2]);

  // Core Reconciliation Engine
  const executeReconciliation = (
    s1: SheetRow[],
    s2: SheetRow[],
    m1Name: string,
    m2Name: string,
    accName: string,
    repNameCol: string,
    fallbackVal = emptyRepFallback
  ) => {
    const repMap = new Map<string, Array<{ account: string; name: string; isMissingRepName: boolean }>>();
    const irregularsList: IrregularAccountItem[] = [];

    for (let i = 0; i < s2.length; i++) {
      const row = s2[i];
      const rawM = row[m2Name];
      if (rawM !== undefined && rawM !== null && rawM !== '') {
        const key = String(rawM).trim();
        const rawRepName = row[repNameCol];
        const isRepNameEmpty =
          rawRepName === undefined ||
          rawRepName === null ||
          String(rawRepName).trim() === '' ||
          String(rawRepName).trim() === '-' ||
          String(rawRepName).trim() === 'null';

        const repName = isRepNameEmpty ? fallbackVal : String(rawRepName).trim();
        const repAcc =
          accName && row[accName] !== undefined && row[accName] !== ''
            ? String(row[accName]).trim()
            : 'غير متوفر';

        // Audit irregular account numbers
        if (repAcc !== 'غير متوفر') {
          const hasSpecialChars = /[^0-9\-_]/i.test(repAcc);
          const hasSpaces = /\s/.test(repAcc);
          const isUnusuallyShort = repAcc.length < 2;
          if (hasSpecialChars || hasSpaces || isUnusuallyShort) {
            irregularsList.push({
              machine: key,
              account: repAcc,
              repName,
              reason: hasSpaces
                ? 'يحتوي على مسافات'
                : hasSpecialChars
                ? 'يحتوي على رموز/حروف'
                : 'قصير جداً',
            });
          }
        }

        if (!repMap.has(key)) {
          repMap.set(key, []);
        }
        const arr = repMap.get(key)!;
        arr.push({ account: repAcc, name: repName, isMissingRepName: isRepNameEmpty });
      }
    }

    // Audit duplicates in Sheet 1
    const machineCounts = new Map<string, number>();
    for (const row of s1) {
      const m = row[m1Name];
      if (m !== undefined && m !== null && m !== '') {
        const key = String(m).trim();
        machineCounts.set(key, (machineCounts.get(key) || 0) + 1);
      }
    }
    const duplicateList: DuplicateAuditItem[] = [];
    machineCounts.forEach((count, key) => {
      if (count > 1) {
        duplicateList.push({ machine: key, occurrences: count });
      }
    });

    // Generate sequential independent rows
    const summaries: MachineSummary[] = [];
    const flatRows: ExpandedRow[] = [];
    let rowCounter = 0;

    for (let i = 0; i < s1.length; i++) {
      const row = s1[i];
      const mVal = row[m1Name];
      if (mVal === undefined || mVal === null || mVal === '') continue;

      const machineId = String(mVal).trim();
      const reps = repMap.get(machineId) || [];
      const count = reps.length;

      let status: 'single' | 'multi' | 'none' = 'none';
      if (count === 1) status = 'single';
      else if (count > 1) status = 'multi';

      summaries.push({
        index: i + 1,
        machine: machineId,
        repCount: count,
        reps,
        status,
      });

      if (count === 0) {
        rowCounter++;
        flatRows.push({
          id: `row-${rowCounter}`,
          machine: machineId,
          account: 'غير متوفر',
          repName: fallbackVal,
          repOrder: 0,
          totalRepsForMachine: 0,
          status: 'none',
          machineGroupIndex: i + 1,
          isFirstOfGroup: true,
          isLastOfGroup: true,
          isMissingRepName: true,
        });
      } else {
        for (let rIdx = 0; rIdx < count; rIdx++) {
          rowCounter++;
          const r = reps[rIdx];
          flatRows.push({
            id: `row-${rowCounter}`,
            machine: machineId,
            account: r.account,
            repName: r.name,
            repOrder: rIdx + 1,
            totalRepsForMachine: count,
            status,
            machineGroupIndex: i + 1,
            isFirstOfGroup: rIdx === 0,
            isLastOfGroup: rIdx === count - 1,
            isMissingRepName: r.isMissingRepName,
          });
        }
      }
    }

    setMachinesResults(summaries);
    setExpandedRows(flatRows);
    setDuplicates(duplicateList);
    setIrregulars(irregularsList);
    setSelectedMachines(new Set());
    setCurrentPage(1);
  };

  // Start reconciliation from user action
  const runReconciliation = () => {
    if (!sheet1.length || !sheet2.length) {
      showToast('⚠️ يرجى رفع شيت الماكينات وشيت المناديب أولاً، أو اضغط على «عينة فورية».');
      return;
    }

    if (!colM1 || !colM2 || !colRep) {
      showToast('⚠️ يرجى ضبط أعمدة الماكينة والمندوب في التبويب 3.');
      setCurrentTab(3);
      return;
    }

    setProgressState({
      isOpen: true,
      title: 'مطابقة',
      fileName: 'محرك المطابقة المليوني O(1)',
      subtitle: 'جاري مطابقة الماكينات وتفصيل أسطر المناديب وحساباتهم...',
      percent: 30,
      stepText: 'تحليل مصفوفة الهاش ومطابقة شيت المكن بشيت الربط...',
      countText: '',
      isComplete: false,
    });

    setTimeout(() => {
      setProgressState((p) => ({
        ...p,
        percent: 75,
        stepText: 'توليد أسطر التعدد والتحقق من حسابات المناديب...',
      }));

      setTimeout(() => {
        executeReconciliation(sheet1, sheet2, colM1, colM2, colAcc, colRep, emptyRepFallback);
        setProgressState((p) => ({
          ...p,
          percent: 100,
          stepText: 'تمت مطابقة وتفصيل كافة الأسطر بنجاح!',
          isComplete: true,
        }));

        setTimeout(() => {
          setProgressState((p) => ({ ...p, isOpen: false }));
          showToast('⚡ تم تحديث ومطابقة البيانات بنجاح!');
          setCurrentTab(4);
        }, 500);
      }, 350);
    }, 250);
  };

  // Handle Sheet 1 upload
  const handleSheet1Upload = async (file: File) => {
    setProgressState({
      isOpen: true,
      title: 'قراءة',
      fileName: file.name,
      subtitle: 'جاري تحليل خلايا وسجلات ملف الإكسل...',
      percent: 25,
      stepText: 'تحويل ورقة العمل إلى هيكل بيانات JSON...',
      countText: '',
      isComplete: false,
    });

    try {
      setTimeout(async () => {
        setProgressState((p) => ({
          ...p,
          percent: 60,
          stepText: 'استخراج أرقام الماكينات وتدقيق التكرارات...',
        }));

        const data = await parseExcelFile(file);
        setSheet1(data);
        setSheet1FileName(file.name);

        setProgressState((p) => ({
          ...p,
          percent: 100,
          stepText: `تم قراءة ${data.length} ماكينة بنجاح!`,
          isComplete: true,
        }));

        setTimeout(() => {
          setProgressState((p) => ({ ...p, isOpen: false }));
          showToast(`✅ تم قراءة شيت الماكينات (${data.length} ماكينة).`);
          if (sheet2.length > 0 && colM1 && colM2 && colRep) {
            executeReconciliation(data, sheet2, colM1, colM2, colAcc, colRep, emptyRepFallback);
          }
        }, 400);
      }, 300);
    } catch (err) {
      setProgressState((p) => ({ ...p, isOpen: false }));
      showToast('❌ تعذر قراءة الملف، يرجى التأكد من صيغة Excel أو CSV.');
    }
  };

  // Handle Sheet 2 upload
  const handleSheet2Upload = async (file: File) => {
    setProgressState({
      isOpen: true,
      title: 'قراءة',
      fileName: file.name,
      subtitle: 'جاري تحليل خلايا وسجلات ملف الإكسل...',
      percent: 25,
      stepText: 'تحويل ورقة العمل إلى هيكل بيانات JSON...',
      countText: '',
      isComplete: false,
    });

    try {
      setTimeout(async () => {
        setProgressState((p) => ({
          ...p,
          percent: 60,
          stepText: 'استخراج حسابات المناديب وربطها بالماكينات...',
        }));

        const data = await parseExcelFile(file);
        setSheet2(data);
        setSheet2FileName(file.name);

        setProgressState((p) => ({
          ...p,
          percent: 100,
          stepText: `تم قراءة ${data.length} سجل مناديب بنجاح!`,
          isComplete: true,
        }));

        setTimeout(() => {
          setProgressState((p) => ({ ...p, isOpen: false }));
          showToast(`✅ تم قراءة شيت المناديب والحسابات (${data.length} سجل).`);
          if (sheet1.length > 0 && colM1 && colM2 && colRep) {
            executeReconciliation(sheet1, data, colM1, colM2, colAcc, colRep, emptyRepFallback);
          }
        }, 400);
      }, 300);
    } catch (err) {
      setProgressState((p) => ({ ...p, isOpen: false }));
      showToast('❌ تعذر قراءة الملف، يرجى التأكد من صيغة Excel أو CSV.');
    }
  };

  // Instant Sample Demo - Instant, smooth animation!
  const handleLoadSample = () => {
    setProgressState({
      isOpen: true,
      title: 'قراءة',
      fileName: 'نموذج_شيت_ربط_المناديب_والحسابات.csv',
      subtitle: 'جاري تحليل خلايا وسجلات ملف الإكسل...',
      percent: 60,
      stepText: 'تحويل ورقة العمل إلى هيكل بيانات JSON...',
      countText: '',
      isComplete: false,
    });

    setTimeout(() => {
      const sample1: SheetRow[] = [
        { 'رقم الماكينة': 'POS-100', 'الفرع / المنطقة': 'القاهرة - المعادي' },
        { 'رقم الماكينة': 'POS-200', 'الفرع / المنطقة': 'الجيزة - الدقي' },
        { 'رقم الماكينة': 'POS-300', 'الفرع / المنطقة': 'الإسكندرية - سموحة' },
        { 'رقم الماكينة': 'POS-400', 'الفرع / المنطقة': 'المنصورة - المشاية' },
        { 'رقم الماكينة': 'POS-500', 'الفرع / المنطقة': 'طنطا - المحطة' },
        { 'رقم الماكينة': 'POS-600', 'الفرع / المنطقة': 'أسيوط - الجمهورية' },
        { 'رقم الماكينة': 'POS-100', 'الفرع / المنطقة': 'القاهرة - المعادي (تكرار)' },
      ];

      const sample2: SheetRow[] = [
        { 'رقم الماكينة': 'POS-100', 'رقم حساب المندوب': '1234', 'اسم المندوب': 'أحمد محمود سالم' },
        { 'رقم الماكينة': 'POS-100', 'رقم حساب المندوب': '456', 'اسم المندوب': 'محمود حسن رضوان' },
        { 'رقم الماكينة': 'POS-200', 'رقم حساب المندوب': '555', 'اسم المندوب': '' },
        { 'رقم الماكينة': 'POS-300', 'رقم حساب المندوب': '789', 'اسم المندوب': 'خالد عبد الرحمن' },
        { 'رقم الماكينة': 'POS-300', 'رقم حساب المندوب': '890', 'اسم المندوب': 'طارق زياد العتيبي' },
        { 'رقم الماكينة': 'POS-300', 'رقم حساب المندوب': '999', 'اسم المندوب': 'عمر فاروق الشامي' },
        { 'رقم الماكينة': 'POS-400', 'رقم حساب المندوب': '777-ERR', 'اسم المندوب': 'إبراهيم حسني مراد' },
      ];

      setSheet1(sample1);
      setSheet1FileName('عينة_شيت_الماكينات.xlsx');
      setSheet2(sample2);
      setSheet2FileName('عينة_شيت_المناديب_والحسابات.xlsx');

      const m1 = 'رقم الماكينة';
      const m2 = 'رقم الماكينة';
      const acc = 'رقم حساب المندوب';
      const rep = 'اسم المندوب';

      setColM1(m1);
      setColM2(m2);
      setColAcc(acc);
      setColRep(rep);

      executeReconciliation(sample1, sample2, m1, m2, acc, rep, 'لا يوجد مندوب');

      setProgressState((p) => ({
        ...p,
        percent: 100,
        stepText: 'اكتمل تحميل العينة بنجاح!',
        isComplete: true,
      }));

      setTimeout(() => {
        setProgressState((p) => ({ ...p, isOpen: false }));
        setCurrentTab(4);
        showToast('✨ تم تحميل العينة التجريبية وعرض النتائج فوراً!');
      }, 400);
    }, 450);
  };

  // Perform complete reset
  const performResetAll = () => {
    setSheet1([]);
    setSheet2([]);
    setSheet1FileName('');
    setSheet2FileName('');
    setMachinesResults([]);
    setExpandedRows([]);
    setDuplicates([]);
    setIrregulars([]);
    setSelectedMachines(new Set());
    setSearchQuery('');
    setCurrentFilter('all');
    setCurrentPage(1);
    setCurrentTab(1);
    showToast('🗑️ تم مسح كافة البيانات وتفريغ النظام بالكامل.');
  };

  // Export results to Excel
  const handleExportExcel = () => {
    if (!expandedRows.length) {
      showToast('⚠️ لا توجد نتائج لتصديرها. ارفع الشيتات أو اختر «عينة فورية».');
      return;
    }
    exportReconciliationToExcel(expandedRows);
    showToast('📥 تم تصدير ملف الإكسل المفصل بنجاح!');
  };

  // Bulk actions handlers
  const handleToggleSelectRow = (machine: string) => {
    setSelectedMachines((prev) => {
      const next = new Set(prev);
      if (next.has(machine)) next.delete(machine);
      else next.add(machine);
      return next;
    });
  };

  const handleToggleSelectAll = (checked: boolean) => {
    if (checked) {
      const allMs = new Set(displayedSlice.map((r) => r.machine));
      setSelectedMachines(allMs);
    } else {
      setSelectedMachines(new Set());
    }
  };

  const handleBulkReassign = (newRep: string, newAcc: string) => {
    if (selectedMachines.size === 0) return;

    const updatedSheet2 = sheet2.filter((r) => !selectedMachines.has(String(r[colM2]).trim()));
    selectedMachines.forEach((m) => {
      updatedSheet2.push({
        [colM2]: m,
        [colAcc]: newAcc || 'غير متوفر',
        [colRep]: newRep || emptyRepFallback,
      });
    });

    setSheet2(updatedSheet2);
    executeReconciliation(sheet1, updatedSheet2, colM1, colM2, colAcc, colRep, emptyRepFallback);
    setSelectedMachines(new Set());
    showToast(`✅ تم إعادة توزيع ${selectedMachines.size} ماكينة للمندوب بنجاح!`);
  };

  const handleBulkMarkVacant = () => {
    if (selectedMachines.size === 0) return;
    const updatedSheet2 = sheet2.filter((r) => !selectedMachines.has(String(r[colM2]).trim()));
    setSheet2(updatedSheet2);
    executeReconciliation(sheet1, updatedSheet2, colM1, colM2, colAcc, colRep, emptyRepFallback);
    setSelectedMachines(new Set());
    showToast(`✅ تم تحويل ${selectedMachines.size} ماكينة إلى شاغرة (بدون مندوب).`);
  };

  // Filtered rows for Tab 4
  const filteredRows = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();
    return expandedRows.filter((item) => {
      if (currentFilter !== 'all' && item.status !== currentFilter) return false;
      if (q) {
        const matchM = item.machine.toLowerCase().includes(q);
        const matchAcc = item.account.toLowerCase().includes(q);
        const matchRep = item.repName.toLowerCase().includes(q);
        return matchM || matchAcc || matchRep;
      }
      return true;
    });
  }, [expandedRows, currentFilter, searchQuery]);

  // Paginated slice
  const maxPages = Math.max(1, Math.ceil(filteredRows.length / pageSize));
  const displayedSlice = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return filteredRows.slice(start, start + pageSize);
  }, [filteredRows, currentPage, pageSize]);

  // Virtualization / Windowing Setup for High Performance Rendering in Tab 4
  const [tableScrollTop, setTableScrollTop] = useState(0);
  const [cardsScrollTop, setCardsScrollTop] = useState(0);
  const tableContainerRef = useRef<HTMLDivElement>(null);
  const cardsContainerRef = useRef<HTMLDivElement>(null);
  
  // Heights and Viewport setup
  const tableRowHeight = 53;
  const cardsRowHeight = 230; // estimated grid row height
  const viewportHeight = 550; // container max height
  
  // Dynamic screen columns count detector to align grid layout with virtualization
  const [colsCount, setColsCount] = useState(4);
  useEffect(() => {
    const handleResize = () => {
      const w = window.innerWidth;
      if (w < 768) setColsCount(1);
      else if (w < 1024) setColsCount(2);
      else if (w < 1280) setColsCount(3);
      else setColsCount(4);
    };
    handleResize();
    window.addEventListener('resize', handleResize);
    return () => window.removeEventListener('resize', handleResize);
  }, []);

  // Reset scroll offsets when current filter, searchQuery, viewMode, or currentPage changes
  useEffect(() => {
    setTableScrollTop(0);
    setCardsScrollTop(0);
    if (tableContainerRef.current) tableContainerRef.current.scrollTop = 0;
    if (cardsContainerRef.current) cardsContainerRef.current.scrollTop = 0;
  }, [currentPage, searchQuery, currentFilter, viewMode]);

  // Memoized table virtualization parameters
  const tableVirtualState = useMemo(() => {
    const visibleCount = Math.ceil(viewportHeight / tableRowHeight);
    const start = Math.max(0, Math.floor(tableScrollTop / tableRowHeight) - 3); // 3 buffer rows
    const end = Math.min(displayedSlice.length, start + visibleCount + 6); // 6 buffer rows
    const topPadding = start * tableRowHeight;
    const bottomPadding = Math.max(0, (displayedSlice.length - end) * tableRowHeight);
    
    return {
      startIndex: start,
      endIndex: end,
      topPadding,
      bottomPadding,
    };
  }, [tableScrollTop, displayedSlice, viewportHeight]);

  // Memoized cards virtualization parameters
  const cardsVirtualState = useMemo(() => {
    const totalRows = Math.ceil(displayedSlice.length / colsCount);
    const visibleRows = Math.ceil(viewportHeight / cardsRowHeight);
    const startRow = Math.max(0, Math.floor(cardsScrollTop / cardsRowHeight) - 1); // 1 buffer row
    const endRow = Math.min(totalRows, startRow + visibleRows + 2); // 2 buffer rows
    
    const startIndex = startRow * colsCount;
    const endIndex = Math.min(displayedSlice.length, endRow * colsCount);
    const topPadding = startRow * cardsRowHeight;
    const bottomPadding = Math.max(0, (totalRows - endRow) * cardsRowHeight);
    
    return {
      startIndex,
      endIndex,
      topPadding,
      bottomPadding,
    };
  }, [cardsScrollTop, displayedSlice, colsCount, viewportHeight]);

  // KPIs
  const singleCount = machinesResults.filter((m) => m.status === 'single').length;
  const multiCount = machinesResults.filter((m) => m.status === 'multi').length;
  const multiRowsCount = expandedRows.filter((r) => r.status === 'multi').length;
  const noneCount = machinesResults.filter((m) => m.status === 'none').length;

  return (
    <div className="min-h-screen flex flex-col bg-slate-50 dark:bg-[#090d16] text-slate-800 dark:text-slate-100 transition-colors duration-300">
      {/* Offline Connectivity Notification */}
      <OfflineIndicator />

      {/* Floating Toast Notification */}
      {toastMessage && (
        <div className="fixed top-20 right-1/2 translate-x-1/2 z-50 px-4 py-2.5 rounded-2xl bg-slate-900/95 text-white border border-indigo-500/40 shadow-2xl text-xs font-bold flex items-center gap-2 animate-in fade-in slide-in-from-top-3 backdrop-blur-md">
          <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0" />
          <span>{toastMessage}</span>
          <button
            onClick={() => setToastMessage(null)}
            className="p-1 text-slate-400 hover:text-white"
          >
            <X className="w-3.5 h-3.5" />
          </button>
        </div>
      )}

      {/* Interactive Global Progress Modal matching image.png */}
      <ProgressModal
        isOpen={progressState.isOpen}
        title={progressState.title}
        fileName={progressState.fileName}
        subtitle={progressState.subtitle}
        percent={progressState.percent}
        stepText={progressState.stepText}
        countText={progressState.countText}
        isComplete={progressState.isComplete}
      />

      {/* Desktop App Center Modal */}
      <DesktopAppCenterModal
        isOpen={isDesktopCenterOpen}
        onClose={() => setIsDesktopCenterOpen(false)}
      />

      {/* Manual Online/Offline Table Editor Modal */}
      <ManualEditorModal
        isOpen={isManualEditorOpen}
        onClose={() => setIsManualEditorOpen(false)}
        initialMode={editorInitialMode}
        currentSheet1={sheet1}
        currentSheet2={sheet2}
        onApplySheet1={(data) => {
          setSheet1(data);
          setSheet1FileName('شيت_الماكينات_المعدل_يدوياً');
          showToast(`✅ تم اعتماد ${data.length} ماكينة من محرر القوالب.`);
          if (sheet2.length > 0) {
            executeReconciliation(data, sheet2, colM1 || 'رقم الماكينة', colM2 || 'رقم الماكينة', colAcc, colRep);
          }
        }}
        onApplySheet2={(data) => {
          setSheet2(data);
          setSheet2FileName('شيت_المناديب_المعدل_يدوياً');
          showToast(`✅ تم اعتماد ${data.length} سجل مناديب من محرر القوالب.`);
          if (sheet1.length > 0) {
            executeReconciliation(sheet1, data, colM1 || 'رقم الماكينة', colM2 || 'رقم الماكينة', colAcc, colRep);
          }
        }}
      />

      {/* Reset Confirmation Modal */}
      <ResetConfirmModal
        isOpen={isResetConfirmOpen}
        onClose={() => setIsResetConfirmOpen(false)}
        onConfirm={performResetAll}
      />

      {/* Rep Machines Lookup Modal */}
      <RepLookupModal
        isOpen={isRepLookupOpen}
        onClose={() => setIsRepLookupOpen(false)}
        machinesResults={machinesResults}
        expandedRows={expandedRows}
        sheet1={sheet1}
        sheet2={sheet2}
        colM1={colM1}
        colM2={colM2}
        colAcc={colAcc}
        colRep={colRep}
        initialRepQuery={repLookupQuery}
        onLoadSample={handleLoadSample}
        onFilterMainTable={(query) => {
          setSearchQuery(query);
          setCurrentFilter('all');
          setCurrentPage(1);
          setCurrentTab(4);
          showToast(`🔍 تم تصفية الجدول للماكينات المربوطة بالمندوب: ${query}`);
        }}
      />

      {/* Desktop Header */}
      <DesktopHeader
        onOpenDesktopCenter={() => setIsDesktopCenterOpen(true)}
        onOpenManualEditor={() => {
          setEditorInitialMode(1);
          setIsManualEditorOpen(true);
        }}
        onOpenRepLookup={() => {
          setRepLookupQuery('');
          setIsRepLookupOpen(true);
        }}
        onLoadSample={handleLoadSample}
        onResetAll={() => setIsResetConfirmOpen(true)}
        isDarkMode={isDarkMode}
        onToggleTheme={handleToggleTheme}
      />

      {/* Main Container */}
      <main className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-5 space-y-6 relative z-10 font-['Cairo']">
        {/* Exact Hero Banner from image.png */}
        <div className="bg-[#0a0f1d] border border-slate-800 rounded-2xl p-4 sm:p-5 flex flex-col xl:flex-row items-center justify-between gap-4 w-full shadow-lg">
          {/* Right Side: File Icon + Title + Certified Badge + Description */}
          <div className="flex items-center gap-4 w-full xl:w-auto">
            {/* File Icon Box */}
            <div className="w-12 h-12 rounded-xl bg-blue-950/60 border border-blue-500/30 text-blue-400 flex items-center justify-center text-xl shrink-0 shadow-sm">
              <FileCheck className="w-6 h-6 text-blue-400" />
            </div>

            <div>
              <div className="flex items-center gap-3 flex-wrap">
                <h2 className="text-lg sm:text-xl font-black text-white tracking-normal">
                  مركز تحميل القوالب الجاهزة والمحرر أوفلاين
                </h2>
                <span className="inline-flex items-center gap-1.5 px-3 py-0.5 rounded-full text-xs font-bold bg-[#062419] border border-emerald-500/40 text-emerald-400">
                  <CheckCircle2 className="w-3.5 h-3.5 text-emerald-400" />
                  <span>رؤوس أعمدة مطابقة 100%</span>
                </span>
              </div>
              <p className="text-xs text-slate-400 mt-1">
                قم بتحميل النماذج الفارغة برؤوس الأعمدة الرسمية، أو استخدم المحرر التفاعلي لإدخال وتعديل البيانات مباشرة.
              </p>
            </div>
          </div>

          {/* Left Side: Buttons in exact order of image.png */}
          <div className="flex flex-wrap items-center gap-2.5 w-full xl:w-auto justify-start xl:justify-end shrink-0">
            {/* 1. محرر القوالب أونلاين Button (Solid Bright Orange/Amber) */}
            <button
              onClick={() => {
                setEditorInitialMode(1);
                setIsManualEditorOpen(true);
              }}
              className="px-5 py-2.5 text-xs font-black bg-[#f59e0b] hover:bg-[#d97706] text-black rounded-xl flex items-center gap-2 shadow-[0_0_20px_rgba(245,158,11,0.35)] transition-all cursor-pointer shrink-0"
            >
              <TableProperties className="w-4 h-4 text-black" />
              <span>محرر القوالب أونلاين</span>
            </button>

            {/* 2. قالب الماكينات (.xlsx) + CSV */}
            <div className="inline-flex items-center bg-[#0d1527] border border-slate-700/80 rounded-xl p-1 text-xs font-bold text-white shrink-0">
              <button
                onClick={() => downloadTemplateFile(1, 'xlsx')}
                className="px-3 py-1.5 flex items-center gap-2 hover:text-emerald-300 transition-colors cursor-pointer"
                title="تنزيل قالب شيت الماكينات إكسل"
              >
                <FileSpreadsheet className="w-4 h-4 text-emerald-400" />
                <span>قالب الماكينات (.xlsx)</span>
              </button>
              <div className="h-4 w-px bg-slate-700 mx-1"></div>
              <button
                onClick={() => downloadTemplateFile(1, 'csv')}
                className="px-2 py-1 text-[11px] font-black text-blue-400 hover:text-blue-300 transition-colors cursor-pointer"
                title="تنزيل بصيغة CSV"
              >
                CSV
              </button>
            </div>

            {/* 3. قالب المناديب (.xlsx) + CSV */}
            <div className="inline-flex items-center bg-[#0d1527] border border-slate-700/80 rounded-xl p-1 text-xs font-bold text-white shrink-0">
              <button
                onClick={() => downloadTemplateFile(2, 'xlsx')}
                className="px-3 py-1.5 flex items-center gap-2 hover:text-emerald-300 transition-colors cursor-pointer"
                title="تنزيل قالب شيت المناديب إكسل"
              >
                <FileSpreadsheet className="w-4 h-4 text-emerald-400" />
                <span>قالب المناديب (.xlsx)</span>
              </button>
              <div className="h-4 w-px bg-slate-700 mx-1"></div>
              <button
                onClick={() => downloadTemplateFile(2, 'csv')}
                className="px-2 py-1 text-[11px] font-black text-blue-400 hover:text-blue-300 transition-colors cursor-pointer"
                title="تنزيل بصيغة CSV"
              >
                CSV
              </button>
            </div>
          </div>
        </div>

        {/* Info & Status Cards */}
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
          <div className="lg:col-span-8 bg-gradient-to-br from-slate-900 via-indigo-950 to-slate-950 text-white rounded-3xl p-6 shadow-xl border border-slate-800 flex flex-col justify-between">
            <div>
              <div className="flex flex-wrap items-center gap-2 mb-2">
                <span className="px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-indigo-500/20 text-indigo-300 border border-indigo-500/30">
                  <Zap className="w-3 h-3 inline ml-1 text-amber-300" /> مطابقة تفصيلية سريعة O(1)
                </span>
                <span className="px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-amber-500/20 text-amber-300 border border-amber-500/30">
                  <Layers className="w-3 h-3 inline ml-1 text-xs" /> سطر مستقل لكل مندوب وحسابه
                </span>
                <span className="px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">
                  <Monitor className="w-3 h-3 inline ml-1 text-xs" /> برنامج ديسك توب أوفلاين
                </span>
              </div>
              <h2 className="text-xl sm:text-2xl font-black mb-2">
                مطابقة ماكينات شيت المكن بحسابات ومناديب شيت الربط
              </h2>
              <p className="text-slate-300 text-xs sm:text-sm leading-relaxed">
                يقوم البرنامج بمطابقة كل ماكينة، وإذا كانت مربوطة بأكثر من مندوب؛ تُعرض{' '}
                <strong className="text-amber-300">تحت بعضها في أسطر متتالية</strong>: رقم الماكينة
                وبجانبها المندوب الأول ورقمه، والسطر التالي تحته مباشرة نفس رقم الماكينة والمندوب الثاني
                ورقمه، مع إمكانية التصدير المباشر لملف إكسل بدون اتصال بالإنترنت.
              </p>
            </div>
            <div className="grid grid-cols-3 gap-2 pt-4 mt-4 border-t border-slate-800/80 text-xs">
              <div className="flex items-center gap-2">
                <span className="w-5 h-5 rounded-md bg-emerald-500/20 text-emerald-400 flex items-center justify-center font-bold">
                  1
                </span>
                <span>مندوب واحد (سطر فردي)</span>
              </div>
              <div className="flex items-center gap-2">
                <span className="w-5 h-5 rounded-md bg-amber-500/20 text-amber-400 flex items-center justify-center font-bold">
                  2+
                </span>
                <span>متعددة المناديب (أسطر متتالية)</span>
              </div>
              <div className="flex items-center gap-2">
                <span className="w-5 h-5 rounded-md bg-rose-500/20 text-rose-400 flex items-center justify-center font-bold">
                  0
                </span>
                <span>بدون مندوب (شاغرة)</span>
              </div>
            </div>
          </div>

          <div className="lg:col-span-4 bg-[#0a0f1d] border border-slate-800 rounded-3xl p-5 shadow-xl flex flex-col justify-between">
            <div className="space-y-3">
              <h3 className="font-black text-white text-xs flex items-center justify-between">
                <span>حالة الملفات المحملة</span>
                <span
                  className={`text-[10px] font-bold px-2.5 py-0.5 rounded-md border ${
                    sheet1.length && sheet2.length
                      ? 'text-emerald-400 bg-[#08201a] border-emerald-500/30'
                      : 'text-amber-400 bg-amber-950/60 border-amber-500/30'
                  }`}
                >
                  {sheet1.length && sheet2.length ? 'جاهز للمطابقة' : 'بانتظار الملفات'}
                </span>
              </h3>

              {/* Item 1: Sheet 1 */}
              <div className="flex items-center justify-between p-3.5 rounded-2xl bg-[#0e1628] border border-slate-800/90 shadow-sm">
                <div className="flex items-center gap-3">
                  <Calculator className="w-5 h-5 text-blue-400 shrink-0" />
                  <div>
                    <span className="text-xs font-black text-white block">1. شيت الماكينات</span>
                    <span
                      className={`text-[11px] block font-mono mt-0.5 max-w-[170px] truncate ${
                        sheet1.length ? 'text-slate-300 font-semibold' : 'text-slate-500'
                      }`}
                      dir={sheet1.length ? 'ltr' : 'rtl'}
                    >
                      {sheet1.length ? (sheet1FileName || 'شيت_الماكينات.xlsx') : 'لم يتم الرفع'}
                    </span>
                  </div>
                </div>
                {sheet1.length > 0 ? (
                  <span className="text-[11px] font-mono font-bold text-emerald-400 bg-[#08201a] border border-emerald-500/30 px-2.5 py-1 rounded-lg">
                    تم ({sheet1.length.toLocaleString()})
                  </span>
                ) : (
                  <span className="text-[11px] font-bold text-rose-400 bg-[#200d14] border border-rose-500/30 px-2.5 py-1 rounded-lg">
                    غير مرفوع
                  </span>
                )}
              </div>

              {/* Item 2: Sheet 2 */}
              <div className="flex items-center justify-between p-3.5 rounded-2xl bg-[#0e1628] border border-slate-800/90 shadow-sm">
                <div className="flex items-center gap-3">
                  <Users className="w-5 h-5 text-blue-400 shrink-0" />
                  <div>
                    <span className="text-xs font-black text-white block">2. شيت المناديب والحسابات</span>
                    <span
                      className={`text-[11px] block font-mono mt-0.5 max-w-[170px] truncate ${
                        sheet2.length ? 'text-slate-300 font-semibold' : 'text-slate-500'
                      }`}
                      dir={sheet2.length ? 'ltr' : 'rtl'}
                    >
                      {sheet2.length ? (sheet2FileName || 'شيت_المناديب.xlsx') : 'لم يتم الرفع'}
                    </span>
                  </div>
                </div>
                {sheet2.length > 0 ? (
                  <span className="text-[11px] font-mono font-bold text-emerald-400 bg-[#08201a] border border-emerald-500/30 px-2.5 py-1 rounded-lg">
                    تم ({sheet2.length.toLocaleString()})
                  </span>
                ) : (
                  <span className="text-[11px] font-bold text-rose-400 bg-[#200d14] border border-rose-500/30 px-2.5 py-1 rounded-lg">
                    غير مرفوع
                  </span>
                )}
              </div>
            </div>

            <button
              onClick={runReconciliation}
              className="mt-4 w-full bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-black py-3 px-4 rounded-2xl shadow-lg shadow-indigo-600/30 transition-all flex items-center justify-center gap-2 cursor-pointer"
            >
              <Zap className="w-4 h-4 text-amber-300" />
              <span>بدء مطابقة وتفصيل أسطر المناديب O(1)</span>
            </button>
          </div>
        </div>

        {/* Navigation Tabs - Modern Cards Grid */}
        <div className="grid grid-cols-2 sm:grid-cols-3 xl:grid-cols-6 gap-2.5 sm:gap-3">
          {/* Card 1: Sheet 1 */}
          <button
            onClick={() => setCurrentTab(1)}
            className={`p-3 sm:p-3.5 rounded-2xl border text-right transition-all duration-200 flex flex-col justify-between gap-2.5 cursor-pointer group relative overflow-hidden ${
              currentTab === 1
                ? 'bg-indigo-600 border-indigo-400 text-white shadow-lg shadow-indigo-600/30 ring-2 ring-indigo-400/40'
                : 'bg-white dark:bg-[#0a0f1d] border-slate-200 dark:border-slate-800 text-slate-700 dark:text-slate-300 hover:border-indigo-400/60 hover:bg-slate-50 dark:hover:bg-[#0e1628]'
            }`}
          >
            <div className="flex items-center justify-between w-full">
              <span
                className={`text-[10px] font-mono font-black px-2 py-0.5 rounded-md ${
                  currentTab === 1
                    ? 'bg-white/20 text-white'
                    : 'bg-slate-100 dark:bg-slate-800 text-slate-500 dark:text-slate-400'
                }`}
              >
                01
              </span>
              <div
                className={`w-7 h-7 rounded-xl flex items-center justify-center transition-colors ${
                  currentTab === 1
                    ? 'bg-white/20 text-white'
                    : 'bg-indigo-500/10 text-indigo-500 dark:text-indigo-400 group-hover:bg-indigo-500/20'
                }`}
              >
                <UploadCloud className="w-4 h-4" />
              </div>
            </div>

            <div>
              <h4
                className={`text-xs font-bold leading-tight ${
                  currentTab === 1 ? 'text-white' : 'text-slate-900 dark:text-white'
                }`}
              >
                1. شيت الماكينات
              </h4>
              <p
                className={`text-[10px] mt-1 font-medium truncate ${
                  currentTab === 1 ? 'text-indigo-100 font-bold' : 'text-slate-500 dark:text-slate-400'
                }`}
              >
                {sheet1.length > 0 ? `${sheet1.length.toLocaleString('ar-EG')} ماكينة` : 'رفع ملف الأرقام'}
              </p>
            </div>

            {currentTab === 1 && (
              <div className="absolute bottom-0 left-0 right-0 h-1 bg-white/70 rounded-full mx-3 mb-0.5" />
            )}
          </button>

          {/* Card 2: Sheet 2 */}
          <button
            onClick={() => setCurrentTab(2)}
            className={`p-3 sm:p-3.5 rounded-2xl border text-right transition-all duration-200 flex flex-col justify-between gap-2.5 cursor-pointer group relative overflow-hidden ${
              currentTab === 2
                ? 'bg-indigo-600 border-indigo-400 text-white shadow-lg shadow-indigo-600/30 ring-2 ring-indigo-400/40'
                : 'bg-white dark:bg-[#0a0f1d] border-slate-200 dark:border-slate-800 text-slate-700 dark:text-slate-300 hover:border-indigo-400/60 hover:bg-slate-50 dark:hover:bg-[#0e1628]'
            }`}
          >
            <div className="flex items-center justify-between w-full">
              <span
                className={`text-[10px] font-mono font-black px-2 py-0.5 rounded-md ${
                  currentTab === 2
                    ? 'bg-white/20 text-white'
                    : 'bg-slate-100 dark:bg-slate-800 text-slate-500 dark:text-slate-400'
                }`}
              >
                02
              </span>
              <div
                className={`w-7 h-7 rounded-xl flex items-center justify-center transition-colors ${
                  currentTab === 2
                    ? 'bg-white/20 text-white'
                    : 'bg-blue-500/10 text-blue-500 dark:text-blue-400 group-hover:bg-blue-500/20'
                }`}
              >
                <Users className="w-4 h-4" />
              </div>
            </div>

            <div>
              <h4
                className={`text-xs font-bold leading-tight ${
                  currentTab === 2 ? 'text-white' : 'text-slate-900 dark:text-white'
                }`}
              >
                2. شيت المناديب
              </h4>
              <p
                className={`text-[10px] mt-1 font-medium truncate ${
                  currentTab === 2 ? 'text-indigo-100 font-bold' : 'text-slate-500 dark:text-slate-400'
                }`}
              >
                {sheet2.length > 0 ? `${sheet2.length.toLocaleString('ar-EG')} سجل توزيع` : 'ربط الحسابات'}
              </p>
            </div>

            {currentTab === 2 && (
              <div className="absolute bottom-0 left-0 right-0 h-1 bg-white/70 rounded-full mx-3 mb-0.5" />
            )}
          </button>

          {/* Card 3: Mapping */}
          <button
            onClick={() => setCurrentTab(3)}
            className={`p-3 sm:p-3.5 rounded-2xl border text-right transition-all duration-200 flex flex-col justify-between gap-2.5 cursor-pointer group relative overflow-hidden ${
              currentTab === 3
                ? 'bg-indigo-600 border-indigo-400 text-white shadow-lg shadow-indigo-600/30 ring-2 ring-indigo-400/40'
                : 'bg-white dark:bg-[#0a0f1d] border-slate-200 dark:border-slate-800 text-slate-700 dark:text-slate-300 hover:border-indigo-400/60 hover:bg-slate-50 dark:hover:bg-[#0e1628]'
            }`}
          >
            <div className="flex items-center justify-between w-full">
              <span
                className={`text-[10px] font-mono font-black px-2 py-0.5 rounded-md ${
                  currentTab === 3
                    ? 'bg-white/20 text-white'
                    : 'bg-slate-100 dark:bg-slate-800 text-slate-500 dark:text-slate-400'
                }`}
              >
                03
              </span>
              <div
                className={`w-7 h-7 rounded-xl flex items-center justify-center transition-colors ${
                  currentTab === 3
                    ? 'bg-white/20 text-white'
                    : 'bg-amber-500/10 text-amber-500 dark:text-amber-400 group-hover:bg-amber-500/20'
                }`}
              >
                <Sliders className="w-4 h-4" />
              </div>
            </div>

            <div>
              <h4
                className={`text-xs font-bold leading-tight ${
                  currentTab === 3 ? 'text-white' : 'text-slate-900 dark:text-white'
                }`}
              >
                3. تعيين الأعمدة
              </h4>
              <p
                className={`text-[10px] mt-1 font-medium truncate ${
                  currentTab === 3 ? 'text-indigo-100 font-bold' : 'text-slate-500 dark:text-slate-400'
                }`}
              >
                {colM1 && colM2 ? 'أعمدة الربط محددة' : 'ضبط ومطابقة الأعمدة'}
              </p>
            </div>

            {currentTab === 3 && (
              <div className="absolute bottom-0 left-0 right-0 h-1 bg-white/70 rounded-full mx-3 mb-0.5" />
            )}
          </button>

          {/* Card 4: Reports */}
          <button
            onClick={() => setCurrentTab(4)}
            className={`p-3 sm:p-3.5 rounded-2xl border text-right transition-all duration-200 flex flex-col justify-between gap-2.5 cursor-pointer group relative overflow-hidden ${
              currentTab === 4
                ? 'bg-indigo-600 border-indigo-400 text-white shadow-lg shadow-indigo-600/30 ring-2 ring-indigo-400/40'
                : 'bg-white dark:bg-[#0a0f1d] border-slate-200 dark:border-slate-800 text-slate-700 dark:text-slate-300 hover:border-indigo-400/60 hover:bg-slate-50 dark:hover:bg-[#0e1628]'
            }`}
          >
            <div className="flex items-center justify-between w-full">
              <span
                className={`text-[10px] font-mono font-black px-2 py-0.5 rounded-md ${
                  currentTab === 4
                    ? 'bg-white/20 text-white'
                    : 'bg-slate-100 dark:bg-slate-800 text-slate-500 dark:text-slate-400'
                }`}
              >
                04
              </span>
              <div
                className={`w-7 h-7 rounded-xl flex items-center justify-center transition-colors ${
                  currentTab === 4
                    ? 'bg-white/20 text-white'
                    : 'bg-indigo-500/10 text-indigo-500 dark:text-indigo-400 group-hover:bg-indigo-500/20'
                }`}
              >
                <Layers className="w-4 h-4" />
              </div>
            </div>

            <div>
              <h4
                className={`text-xs font-bold leading-tight ${
                  currentTab === 4 ? 'text-white' : 'text-slate-900 dark:text-white'
                }`}
              >
                4. تقرير الأسطر
              </h4>
              <p
                className={`text-[10px] mt-1 font-mono font-bold truncate ${
                  currentTab === 4 ? 'text-indigo-100' : 'text-indigo-600 dark:text-indigo-400'
                }`}
              >
                {expandedRows.length.toLocaleString('ar-EG')} سطر مفصل
              </p>
            </div>

            {currentTab === 4 && (
              <div className="absolute bottom-0 left-0 right-0 h-1 bg-white/70 rounded-full mx-3 mb-0.5" />
            )}
          </button>

          {/* Card 5: Audit & Conflicts */}
          <button
            onClick={() => setCurrentTab(5)}
            className={`p-3 sm:p-3.5 rounded-2xl border text-right transition-all duration-200 flex flex-col justify-between gap-2.5 cursor-pointer group relative overflow-hidden ${
              currentTab === 5
                ? 'bg-amber-600 border-amber-400 text-white shadow-lg shadow-amber-600/30 ring-2 ring-amber-400/40'
                : 'bg-white dark:bg-[#0a0f1d] border-slate-200 dark:border-slate-800 text-slate-700 dark:text-slate-300 hover:border-amber-400/60 hover:bg-slate-50 dark:hover:bg-[#0e1628]'
            }`}
          >
            <div className="flex items-center justify-between w-full">
              <span
                className={`text-[10px] font-mono font-black px-2 py-0.5 rounded-md ${
                  currentTab === 5
                    ? 'bg-white/20 text-white'
                    : 'bg-slate-100 dark:bg-slate-800 text-slate-500 dark:text-slate-400'
                }`}
              >
                05
              </span>
              <div
                className={`w-7 h-7 rounded-xl flex items-center justify-center transition-colors ${
                  currentTab === 5
                    ? 'bg-white/20 text-white'
                    : 'bg-amber-500/10 text-amber-500 dark:text-amber-400 group-hover:bg-amber-500/20'
                }`}
              >
                <CircleAlert className="w-4 h-4" />
              </div>
            </div>

            <div>
              <h4
                className={`text-xs font-bold leading-tight ${
                  currentTab === 5 ? 'text-white' : 'text-amber-600 dark:text-amber-400'
                }`}
              >
                5. فحص النزاعات
              </h4>
              <p
                className={`text-[10px] mt-1 font-mono font-bold truncate ${
                  currentTab === 5 ? 'text-amber-100' : 'text-slate-500 dark:text-slate-400'
                }`}
              >
                {duplicates.length + irregulars.length > 0
                  ? `${(duplicates.length + irregulars.length).toLocaleString('ar-EG')} شاذ / مكرر`
                  : 'فحص سليم (0)'}
              </p>
            </div>

            {currentTab === 5 && (
              <div className="absolute bottom-0 left-0 right-0 h-1 bg-white/70 rounded-full mx-3 mb-0.5" />
            )}
          </button>

          {/* Card 6: Desktop App Studio */}
          <button
            onClick={() => setCurrentTab(6)}
            className={`p-3 sm:p-3.5 rounded-2xl border text-right transition-all duration-200 flex flex-col justify-between gap-2.5 cursor-pointer group relative overflow-hidden ${
              currentTab === 6
                ? 'bg-blue-600 border-blue-400 text-white shadow-lg shadow-blue-600/30 ring-2 ring-blue-400/40'
                : 'bg-white dark:bg-[#0a0f1d] border-slate-200 dark:border-slate-800 text-slate-700 dark:text-slate-300 hover:border-blue-400/60 hover:bg-slate-50 dark:hover:bg-[#0e1628]'
            }`}
          >
            <div className="flex items-center justify-between w-full">
              <span
                className={`text-[10px] font-mono font-black px-2 py-0.5 rounded-md ${
                  currentTab === 6
                    ? 'bg-white/20 text-white'
                    : 'bg-slate-100 dark:bg-slate-800 text-slate-500 dark:text-slate-400'
                }`}
              >
                06
              </span>
              <div
                className={`w-7 h-7 rounded-xl flex items-center justify-center transition-colors ${
                  currentTab === 6
                    ? 'bg-white/20 text-white'
                    : 'bg-blue-500/10 text-blue-500 dark:text-blue-400 group-hover:bg-blue-500/20'
                }`}
              >
                <FileCode className="w-4 h-4" />
              </div>
            </div>

            <div>
              <div className="flex items-center justify-between gap-1">
                <h4
                  className={`text-xs font-bold leading-tight ${
                    currentTab === 6 ? 'text-white' : 'text-slate-900 dark:text-white'
                  }`}
                >
                  6. استوديو HTML
                </h4>
                <span className="w-2 h-2 rounded-full bg-emerald-400 animate-pulse shrink-0" title="Live Update" />
              </div>
              <p
                className={`text-[10px] mt-1 font-medium truncate ${
                  currentTab === 6 ? 'text-blue-100 font-bold' : 'text-slate-500 dark:text-slate-400'
                }`}
              >
                تحديث أوفلاين
              </p>
            </div>

            {currentTab === 6 && (
              <div className="absolute bottom-0 left-0 right-0 h-1 bg-white/70 rounded-full mx-3 mb-0.5" />
            )}
          </button>
        </div>

        {/* TAB 1: Sheet 1 Upload */}
        {currentTab === 1 && (
          <div className="space-y-4">
            <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-3xl p-6 sm:p-8">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-4">
                <div>
                  <h3 className="text-sm font-bold text-slate-900 dark:text-white mb-1 flex items-center gap-2">
                    <UploadCloud className="w-4 h-4 text-indigo-500" />
                    <span>رفع الشيت الأول (شيت أرقام الماكينات المستهدفة)</span>
                  </h3>
                  <p className="text-xs text-slate-500 dark:text-slate-400">
                    الملف الذي يحتوي على أرقام الماكينات (Machine ID / Serial) المطلوب فحصها ومعرفة مناديبها وحساباتهم.
                  </p>
                </div>

                <div className="flex items-center gap-2">
                  <button
                    onClick={() => {
                      setEditorInitialMode(1);
                      setIsManualEditorOpen(true);
                    }}
                    className="px-3 py-1.5 text-xs font-bold rounded-xl bg-slate-100 hover:bg-slate-200 dark:bg-slate-800 dark:hover:bg-slate-700 text-indigo-600 dark:text-indigo-400 border border-slate-200 dark:border-slate-700 flex items-center gap-1.5 transition-colors cursor-pointer"
                  >
                    <TableProperties className="w-3.5 h-3.5" />
                    <span>تعبئة القالب يدوياً</span>
                  </button>
                  <button
                    onClick={() => downloadTemplateFile(1, 'xlsx')}
                    className="px-3 py-1.5 text-xs font-bold rounded-xl bg-indigo-50 hover:bg-indigo-100 dark:bg-indigo-950/60 text-indigo-600 dark:text-indigo-300 border border-indigo-200 dark:border-indigo-800 flex items-center gap-1.5 transition-colors cursor-pointer"
                  >
                    <Download className="w-3.5 h-3.5" />
                    <span>تحميل نموذج القالب</span>
                  </button>
                </div>
              </div>

              {/* Dropzone */}
              <label className="relative border-2 border-dashed border-slate-300 dark:border-slate-700 rounded-3xl p-8 text-center cursor-pointer hover:border-indigo-500 bg-slate-50 dark:bg-slate-800/30 block transition-colors">
                <input
                  type="file"
                  accept=".xlsx, .xls, .csv"
                  className="hidden"
                  onChange={(e) => {
                    if (e.target.files?.length) {
                      handleSheet1Upload(e.target.files[0]);
                    }
                  }}
                />
                <FileSpreadsheet className="w-12 h-12 text-indigo-500 mx-auto mb-2" />
                <p className="text-xs font-bold text-slate-800 dark:text-slate-200 mb-1">
                  اضغط هنا لاختيار شيت الماكينات أو اسحبه وأفلته هنا
                </p>
                <p className="text-[11px] text-slate-400">
                  يدعم ملفات Excel (.xlsx, .xls) أو CSV مع معالجة محلية 100%
                </p>
              </label>

              {sheet1.length > 0 && (
                <div className="mt-4 p-3 rounded-xl bg-emerald-500/10 border border-emerald-500/20 text-emerald-600 dark:text-emerald-400 text-xs font-bold flex items-center justify-between">
                  <span>تم قراءة شيت الماكينات بنجاح: {sheet1.length} ماكينة</span>
                  <button
                    onClick={() => setCurrentTab(2)}
                    className="text-indigo-600 dark:text-indigo-400 underline font-bold cursor-pointer"
                  >
                    الانتقال لرفع شيت المناديب &larr;
                  </button>
                </div>
              )}
            </div>
          </div>
        )}

        {/* TAB 2: Sheet 2 Upload */}
        {currentTab === 2 && (
          <div className="space-y-4">
            <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-3xl p-6 sm:p-8">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-4">
                <div>
                  <h3 className="text-sm font-bold text-slate-900 dark:text-white mb-1 flex items-center gap-2">
                    <Users className="w-4 h-4 text-blue-500" />
                    <span>رفع الشيت الثاني (شيت المناديب وأرقام حساباتهم وربط الماكينات)</span>
                  </h3>
                  <p className="text-xs text-slate-500 dark:text-slate-400">
                    الملف الذي يحتوي على رقم الماكينة، اسم المندوب، ورقم حساب المندوب (Account No / Code).
                  </p>
                </div>

                <div className="flex items-center gap-2">
                  <button
                    onClick={() => {
                      setEditorInitialMode(2);
                      setIsManualEditorOpen(true);
                    }}
                    className="px-3 py-1.5 text-xs font-bold rounded-xl bg-slate-100 hover:bg-slate-200 dark:bg-slate-800 dark:hover:bg-slate-700 text-blue-600 dark:text-blue-400 border border-slate-200 dark:border-slate-700 flex items-center gap-1.5 transition-colors cursor-pointer"
                  >
                    <TableProperties className="w-3.5 h-3.5" />
                    <span>تعبئة القالب يدوياً</span>
                  </button>
                  <button
                    onClick={() => downloadTemplateFile(2, 'xlsx')}
                    className="px-3 py-1.5 text-xs font-bold rounded-xl bg-blue-50 hover:bg-blue-100 dark:bg-blue-950/60 text-blue-600 dark:text-blue-300 border border-blue-200 dark:border-blue-800 flex items-center gap-1.5 transition-colors cursor-pointer"
                  >
                    <Download className="w-3.5 h-3.5" />
                    <span>تحميل نموذج القالب</span>
                  </button>
                </div>
              </div>

              {/* Dropzone */}
              <label className="relative border-2 border-dashed border-slate-300 dark:border-slate-700 rounded-3xl p-8 text-center cursor-pointer hover:border-blue-500 bg-slate-50 dark:bg-slate-800/30 block transition-colors">
                <input
                  type="file"
                  accept=".xlsx, .xls, .csv"
                  className="hidden"
                  onChange={(e) => {
                    if (e.target.files?.length) {
                      handleSheet2Upload(e.target.files[0]);
                    }
                  }}
                />
                <FileSpreadsheet className="w-12 h-12 text-blue-500 mx-auto mb-2" />
                <p className="text-xs font-bold text-slate-800 dark:text-slate-200 mb-1">
                  اضغط هنا لاختيار شيت المناديب والحسابات أو اسحبه وأفلته هنا
                </p>
                <p className="text-[11px] text-slate-400">يدعم ملفات Excel (.xlsx, .xls) أو CSV</p>
              </label>

              {sheet2.length > 0 && (
                <div className="mt-4 p-3 rounded-xl bg-emerald-500/10 border border-emerald-500/20 text-emerald-600 dark:text-emerald-400 text-xs font-bold flex items-center justify-between">
                  <span>تم قراءة شيت المناديب والحسابات: {sheet2.length} سجل توزيع</span>
                  <button
                    onClick={() => setCurrentTab(3)}
                    className="text-indigo-600 dark:text-indigo-400 underline font-bold cursor-pointer"
                  >
                    الانتقال لضبط الأعمدة &larr;
                  </button>
                </div>
              )}
            </div>
          </div>
        )}



        {/* TAB 3: Column Mapping */}
        {currentTab === 3 && (
          <div className="space-y-4">
            <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-3xl p-6 sm:p-8">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-4">
                <div>
                  <h3 className="text-sm font-bold text-slate-900 dark:text-white mb-1 flex items-center gap-2">
                    <Sliders className="w-4 h-4 text-indigo-500" />
                    <span>تحديد وتعيين أعمدة الماكينة ورقم الحساب واسم المندوب</span>
                  </h3>
                  <p className="text-xs text-slate-500 dark:text-slate-400">
                    يتعرف النظام آلياً على الأعمدة المقابلة ويمكنك تخصيص وتعديل طريقة المطابقة:
                  </p>
                </div>
              </div>

              {/* Selectors Grid */}
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                {/* Machine Col 1 */}
                <div className="p-4 rounded-2xl bg-slate-50 dark:bg-slate-800/40 border border-slate-200 dark:border-slate-800 space-y-2">
                  <label className="block text-xs font-bold text-indigo-600 dark:text-indigo-400">
                    1. عمود الماكينة (شيت المكن):
                  </label>
                  <select
                    className="w-full p-2 bg-white dark:bg-slate-900 border border-slate-300 dark:border-slate-700 rounded-xl text-xs font-semibold focus:border-indigo-500 focus:outline-none"
                    value={colM1}
                    onChange={(e) => setColM1(e.target.value)}
                  >
                    {sheet1.length > 0 ? (
                      Object.keys(sheet1[0]).map((c) => (
                        <option key={c} value={c}>
                          {c}
                        </option>
                      ))
                    ) : (
                      <option value="">-- يرجى رفع شيت الماكينات أولاً --</option>
                    )}
                  </select>
                  <p className="text-[10px] text-slate-400">عمود الماكينة الرئيسي من الشيت الأول</p>
                </div>

                {/* Machine Col 2 */}
                <div className="p-4 rounded-2xl bg-slate-50 dark:bg-slate-800/40 border border-slate-200 dark:border-slate-800 space-y-2">
                  <label className="block text-xs font-bold text-blue-600 dark:text-blue-400">
                    2. عمود الماكينة (شيت المناديب):
                  </label>
                  <select
                    className="w-full p-2 bg-white dark:bg-slate-900 border border-slate-300 dark:border-slate-700 rounded-xl text-xs font-semibold focus:border-blue-500 focus:outline-none"
                    value={colM2}
                    onChange={(e) => setColM2(e.target.value)}
                  >
                    {sheet2.length > 0 ? (
                      Object.keys(sheet2[0]).map((c) => (
                        <option key={c} value={c}>
                          {c}
                        </option>
                      ))
                    ) : (
                      <option value="">-- يرجى رفع شيت المناديب أولاً --</option>
                    )}
                  </select>
                  <p className="text-[10px] text-slate-400">رقم الماكينة المقابل في شيت المناديب</p>
                </div>

                {/* Account Col */}
                <div className="p-4 rounded-2xl bg-slate-50 dark:bg-slate-800/40 border border-slate-200 dark:border-slate-800 space-y-2">
                  <label className="block text-xs font-bold text-amber-600 dark:text-amber-400">
                    3. عمود رقم حساب المندوب:
                  </label>
                  <select
                    className="w-full p-2 bg-white dark:bg-slate-900 border border-slate-300 dark:border-slate-700 rounded-xl text-xs font-semibold focus:border-amber-500 focus:outline-none"
                    value={colAcc}
                    onChange={(e) => setColAcc(e.target.value)}
                  >
                    <option value="">-- بدون عمود حساب --</option>
                    {sheet2.length > 0 &&
                      Object.keys(sheet2[0]).map((c) => (
                        <option key={c} value={c}>
                          {c}
                        </option>
                      ))}
                  </select>
                  <p className="text-[10px] text-slate-400">كود / حساب المندوب أو الوكيل</p>
                </div>

                {/* Rep Name Col */}
                <div className="p-4 rounded-2xl bg-slate-50 dark:bg-slate-800/40 border border-slate-200 dark:border-slate-800 space-y-2">
                  <label className="block text-xs font-bold text-emerald-600 dark:text-emerald-400">
                    4. عمود اسم المندوب:
                  </label>
                  <select
                    className="w-full p-2 bg-white dark:bg-slate-900 border border-slate-300 dark:border-slate-700 rounded-xl text-xs font-semibold focus:border-emerald-500 focus:outline-none"
                    value={colRep}
                    onChange={(e) => setColRep(e.target.value)}
                  >
                    {sheet2.length > 0 ? (
                      Object.keys(sheet2[0]).map((c) => (
                        <option key={c} value={c}>
                          {c}
                        </option>
                      ))
                    ) : (
                      <option value="">-- يرجى رفع شيت المناديب أولاً --</option>
                    )}
                  </select>
                  <p className="text-[10px] text-slate-400">عمود اسم المندوب الفعلي</p>
                </div>
              </div>

              {/* Fallback Option */}
              <div className="mt-4 p-3.5 bg-slate-50 dark:bg-slate-800/60 rounded-2xl border border-slate-200 dark:border-slate-700/80 flex flex-wrap items-center justify-between gap-3 text-xs">
                <div>
                  <span className="font-bold text-slate-800 dark:text-slate-200">
                    معالجة الحالات الشاغرة وأسماء المناديب الفارغة:
                  </span>
                  <span className="text-slate-500 dark:text-slate-400 block sm:inline mr-1 text-[11px]">
                    إذا كانت خانة اسم المندوب فارغة يتم تلقائياً تدوين وسم "لا يوجد مندوب" أو النص الذي تخصصه.
                  </span>
                </div>
                <div className="flex items-center gap-2">
                  <span className="text-[11px] font-bold text-slate-600 dark:text-slate-300">النص البديل:</span>
                  <select
                    className="p-1.5 bg-white dark:bg-slate-900 border border-slate-300 dark:border-slate-700 rounded-lg text-xs font-bold text-rose-600 dark:text-rose-400"
                    value={emptyRepFallback}
                    onChange={(e) => setEmptyRepFallback(e.target.value)}
                  >
                    <option value="لا يوجد مندوب">لا يوجد مندوب</option>
                    <option value="لا يوجد">لا يوجد</option>
                    <option value="بدون مندوب مسند">بدون مندوب مسند</option>
                    <option value="شاغرة">شاغرة</option>
                  </select>
                </div>
              </div>

              <div className="mt-6 pt-4 border-t border-slate-100 dark:border-slate-800 flex justify-between items-center">
                <button
                  onClick={() => setCurrentTab(2)}
                  className="text-xs font-bold text-slate-500 hover:text-slate-700 dark:hover:text-slate-300 flex items-center gap-1.5 cursor-pointer"
                >
                  <ArrowRight className="w-4 h-4" />
                  <span>الرجوع لشيت المناديب</span>
                </button>
                <button
                  onClick={runReconciliation}
                  className="bg-indigo-600 hover:bg-indigo-500 text-white font-bold text-xs py-2.5 px-6 rounded-xl shadow-md flex items-center gap-2 transition-all cursor-pointer"
                >
                  <CheckCheck className="w-4 h-4" />
                  <span>تأكيد الإعدادات وتشغيل المطابقة</span>
                </button>
              </div>
            </div>
          </div>
        )}

        {/* TAB 4: Detailed Reconciliation Reports & Sequential Rows */}
        {currentTab === 4 && (
          <div className="space-y-5">
            {/* KPI Summary Cards */}
            <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3">
              {/* Unique Machines */}
              <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 p-3.5 rounded-2xl shadow-sm">
                <div className="flex items-center justify-between">
                  <span className="text-[11px] text-slate-500 block mb-1">الماكينات الفريدة</span>
                  <Layers className="w-3.5 h-3.5 text-indigo-500" />
                </div>
                <span className="text-xl font-black text-slate-900 dark:text-white">
                  {machinesResults.length.toLocaleString('ar-EG')}
                </span>
                <span className="text-[10px] text-slate-400 block mt-1">
                  ينتج عنها{' '}
                  <strong className="text-indigo-600 dark:text-indigo-400">
                    {expandedRows.length.toLocaleString('ar-EG')}
                  </strong>{' '}
                  سطر
                </span>
              </div>

              {/* Single Rep */}
              <div className="bg-white dark:bg-slate-900 border border-emerald-300 dark:border-emerald-800/60 p-3.5 rounded-2xl shadow-sm">
                <div className="flex items-center justify-between">
                  <span className="text-[11px] text-emerald-600 dark:text-emerald-400 font-bold block mb-1">
                    مندوب واحد
                  </span>
                  <UserCheck className="w-3.5 h-3.5 text-emerald-500" />
                </div>
                <span className="text-xl font-black text-emerald-600 dark:text-emerald-400">
                  {singleCount.toLocaleString('ar-EG')}
                </span>
                <span className="text-[10px] text-slate-400 block mt-1">ماكينة بسطر مستقل</span>
              </div>

              {/* Multi Rep */}
              <div className="bg-white dark:bg-slate-900 border border-amber-300 dark:border-amber-700/80 p-3.5 rounded-2xl bg-amber-500/[0.02] shadow-sm">
                <div className="flex items-center justify-between">
                  <span className="text-[11px] text-amber-600 dark:text-amber-400 font-bold block mb-1">
                    متعددة المناديب
                  </span>
                  <Users className="w-3.5 h-3.5 text-amber-500" />
                </div>
                <span className="text-xl font-black text-amber-600 dark:text-amber-400">
                  {multiCount.toLocaleString('ar-EG')}
                </span>
                <span className="text-[10px] text-amber-600/80 dark:text-amber-400/80 block mt-1">
                  مفصلة في <strong>{multiRowsCount.toLocaleString('ar-EG')}</strong> أسطر
                </span>
              </div>

              {/* Vacant / None */}
              <div className="bg-white dark:bg-slate-900 border border-rose-300 dark:border-rose-800/60 p-3.5 rounded-2xl shadow-sm">
                <div className="flex items-center justify-between">
                  <span className="text-[11px] text-rose-600 dark:text-rose-400 font-bold block mb-1">
                    بدون مندوب
                  </span>
                  <UserX className="w-3.5 h-3.5 text-rose-500" />
                </div>
                <span className="text-xl font-black text-rose-600 dark:text-rose-400">
                  {noneCount.toLocaleString('ar-EG')}
                </span>
                <span className="text-[10px] text-slate-400 block mt-1">ماكينة غير مسندة</span>
              </div>

              {/* Duplicates */}
              <div
                onClick={() => setCurrentTab(5)}
                className="bg-white dark:bg-slate-900 border border-amber-500/40 dark:border-amber-600/40 p-3.5 rounded-2xl cursor-pointer hover:border-amber-500 transition-colors shadow-sm"
              >
                <div className="flex items-center justify-between">
                  <span className="text-[11px] text-amber-600 dark:text-amber-400 font-bold block mb-1">
                    تكرار الشيت 1
                  </span>
                  <Layers className="w-3.5 h-3.5 text-amber-500" />
                </div>
                <span className="text-xl font-black text-amber-500">
                  {duplicates.length.toLocaleString('ar-EG')}
                </span>
                <span className="text-[10px] text-slate-400 block mt-1">ماكينة مكررة</span>
              </div>

              {/* Irregular Accounts */}
              <div
                onClick={() => setCurrentTab(5)}
                className="bg-white dark:bg-slate-900 border border-rose-500/40 dark:border-rose-600/40 p-3.5 rounded-2xl cursor-pointer hover:border-rose-500 transition-colors shadow-sm"
              >
                <div className="flex items-center justify-between">
                  <span className="text-[11px] text-rose-600 dark:text-rose-400 font-bold block mb-1">
                    حسابات شاذة
                  </span>
                  <CircleAlert className="w-3.5 h-3.5 text-rose-500" />
                </div>
                <span className="text-xl font-black text-rose-500">
                  {irregulars.length.toLocaleString('ar-EG')}
                </span>
                <span className="text-[10px] text-slate-400 block mt-1">حساب غير منتظم</span>
              </div>
            </div>

            {/* Table Card */}
            <div className="bg-[#0a0f1d] border border-slate-800 rounded-3xl p-5 space-y-4 shadow-sm font-['Cairo']">
              {/* Notification Alert Banner */}
              <div className="p-3.5 bg-blue-500/10 border border-blue-500/30 rounded-2xl text-xs flex items-center justify-between flex-wrap gap-2.5 shadow-sm">
                <div className="flex items-center gap-2.5 text-slate-200">
                  <div className="w-6 h-6 rounded-full bg-blue-500/20 text-blue-400 flex items-center justify-center shrink-0">
                    <CircleAlert className="w-4 h-4 text-blue-400" />
                  </div>
                  <span className="text-xs text-slate-300 leading-relaxed">
                    يتم تفصيل كل ماكينة متعددة المناديب في أسطر متتالية تحت بعضها (رقم الماكينة ثم رقم حساب المندوب الأول، وتحته مباشرة نفس رقم الماكينة ورقم حساب المندوب الثاني وهكذا).
                  </span>
                </div>
                <span className="text-[11px] font-bold px-3 py-1 rounded-full bg-amber-500/10 text-amber-400 border border-amber-500/20 flex items-center gap-1.5 shrink-0">
                  <CircleAlert className="w-3.5 h-3.5 text-amber-400" />
                  <span>جاهز للعرض والتصدير كسطور منفصلة</span>
                </span>
              </div>

              {/* Toolbar & Filter Controls (Glassmorphism Bar) */}
              <div className="p-3.5 rounded-2xl bg-[#0c1220]/80 backdrop-blur-xl border border-slate-800 flex flex-col lg:flex-row justify-between items-center gap-3.5 shadow-lg">
                {/* Right Side: Search Input & Segmented View Controls */}
                <div className="flex items-center gap-3 w-full lg:w-auto flex-wrap">
                  {/* Search Input */}
                  <div className="relative flex-1 sm:w-80">
                    <Search className="w-4 h-4 absolute right-3.5 top-3 text-slate-400 text-xs pointer-events-none" />
                    <input
                      ref={searchInputRef}
                      type="text"
                      autoComplete="off"
                      autoCorrect="off"
                      spellCheck="false"
                      className="w-full pr-10 pl-9 py-2 bg-slate-900 border border-slate-700/80 rounded-xl text-xs font-mono font-bold text-white focus:outline-none focus:border-indigo-500 shadow-inner"
                      placeholder="بحث برقم حساب المندوب أو الماكينة..."
                      value={searchQuery}
                      onChange={(e) => {
                        setSearchQuery(e.target.value);
                        setCurrentPage(1);
                      }}
                    />
                    {searchQuery && (
                      <button
                        type="button"
                        onClick={() => {
                          setSearchQuery('');
                          setCurrentPage(1);
                          searchInputRef.current?.focus();
                        }}
                        className="absolute left-3 top-1/2 -translate-y-1/2 w-5 h-5 rounded-full bg-slate-800 hover:bg-slate-700 text-slate-400 hover:text-white flex items-center justify-center text-xs transition cursor-pointer"
                        title="مسح البحث"
                      >
                        <X className="w-3 h-3" />
                      </button>
                    )}
                  </div>

                  {/* Segmented Control for Table vs Cards */}
                  <div className="inline-flex rounded-xl p-1 bg-slate-900 border border-slate-800 text-xs font-bold shrink-0">
                    <button
                      onClick={() => setViewMode('table')}
                      className={`px-3 py-1.5 rounded-lg flex items-center gap-1.5 transition-all cursor-pointer ${
                        viewMode === 'table'
                          ? 'bg-indigo-600 text-white shadow-sm'
                          : 'text-slate-400 hover:text-white'
                      }`}
                      title="عرض البيانات كجدول تقليدي"
                    >
                      <TableProperties className="w-3.5 h-3.5" />
                      <span>عرض الجدول (Table)</span>
                    </button>
                    <button
                      onClick={() => setViewMode('cards')}
                      className={`px-3 py-1.5 rounded-lg flex items-center gap-1.5 transition-all cursor-pointer ${
                        viewMode === 'cards'
                          ? 'bg-indigo-600 text-white shadow-sm'
                          : 'text-slate-400 hover:text-white'
                      }`}
                      title="عرض البيانات كبطاقات وبوكسات تفاعلية"
                    >
                      <Layers className="w-3.5 h-3.5" />
                      <span>عرض البطاقات (Cards)</span>
                    </button>
                  </div>
                </div>

                {/* Left Side: Filter Buttons Capsule + Excel Export Button */}
                <div className="flex items-center gap-3 w-full lg:w-auto justify-end flex-wrap">
                  {/* Filter Buttons Capsule */}
                  <div className="inline-flex rounded-xl p-1 bg-slate-900 border border-slate-800 text-xs font-bold shrink-0 flex-wrap">
                    <button
                      onClick={() => {
                        setRepLookupQuery('');
                        setIsRepLookupOpen(true);
                      }}
                      className="px-3 py-1.5 rounded-lg bg-blue-600 text-white shadow-sm flex items-center gap-1.5 cursor-pointer"
                    >
                      <UserCheck className="w-3.5 h-3.5" />
                      <span>استعلام ماكينات مندوب</span>
                    </button>
                    <button
                      onClick={() => {
                        setCurrentFilter('all');
                        setCurrentPage(1);
                      }}
                      className={`px-3 py-1.5 rounded-lg transition-colors cursor-pointer ${
                        currentFilter === 'all'
                          ? 'bg-indigo-600 text-white shadow-sm'
                          : 'text-slate-300 hover:text-white'
                      }`}
                    >
                      الكل ({expandedRows.length || 11})
                    </button>
                    <button
                      onClick={() => {
                        setCurrentFilter('single');
                        setCurrentPage(1);
                      }}
                      className={`px-3 py-1.5 rounded-lg transition-colors cursor-pointer ${
                        currentFilter === 'single'
                          ? 'bg-indigo-600 text-white shadow-sm'
                          : 'text-slate-300 hover:text-white'
                      }`}
                    >
                      مندوب واحد
                    </button>
                    <button
                      onClick={() => {
                        setCurrentFilter('multi');
                        setCurrentPage(1);
                      }}
                      className={`px-3 py-1.5 rounded-lg transition-colors cursor-pointer ${
                        currentFilter === 'multi'
                          ? 'bg-indigo-600 text-white shadow-sm'
                          : 'text-slate-300 hover:text-white'
                      }`}
                    >
                      متعددة المناديب
                    </button>
                    <button
                      onClick={() => {
                        setCurrentFilter('none');
                        setCurrentPage(1);
                      }}
                      className={`px-3 py-1.5 rounded-lg transition-colors cursor-pointer ${
                        currentFilter === 'none'
                          ? 'bg-indigo-600 text-white shadow-sm'
                          : 'text-slate-300 hover:text-white'
                      }`}
                    >
                      شاغرة
                    </button>
                  </div>

                  {/* Excel Export Button */}
                  <button
                    onClick={handleExportExcel}
                    className="px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold flex items-center gap-2 shadow-md transition-colors cursor-pointer shrink-0"
                  >
                    <FileSpreadsheet className="w-4 h-4 text-emerald-200" />
                    <span>تصدير Excel (أسطر منفصلة)</span>
                  </button>
                </div>
              </div>

              {/* View Mode: CARDS OR TABLE */}
              {viewMode === 'cards' ? (
                <div className="space-y-4">
                  {displayedSlice.length === 0 ? (
                    <div className="p-12 text-center text-slate-400 bg-slate-50 dark:bg-slate-800/40 rounded-3xl border border-slate-200 dark:border-slate-800 space-y-2">
                      <CircleAlert className="w-8 h-8 text-amber-500 mx-auto" />
                      <p className="font-bold text-sm">لا توجد نتائج تطابق شرط البحث أو الفلتر.</p>
                      <p className="text-xs">ارفع الشيتات أو اختر «عينة فورية» لتوليد البطاقات تلقائياً.</p>
                    </div>
                  ) : (
                    <div
                      ref={cardsContainerRef}
                      onScroll={(e) => setCardsScrollTop(e.currentTarget.scrollTop)}
                      className="max-h-[550px] overflow-y-auto pr-1"
                    >
                      <div
                        style={{
                          paddingTop: cardsVirtualState.topPadding,
                          paddingBottom: cardsVirtualState.bottomPadding,
                        }}
                      >
                        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-4">
                          {displayedSlice.slice(cardsVirtualState.startIndex, cardsVirtualState.endIndex).map((item, idx) => {
                            const globalIdx = (currentPage - 1) * pageSize + cardsVirtualState.startIndex + idx + 1;
                            const isSelected = selectedMachines.has(item.machine);

                            return (
                              <div
                                key={item.id}
                                className={`p-4 rounded-3xl border-2 transition-all shadow-md flex flex-col justify-between gap-3 relative overflow-hidden group ${
                                  item.status === 'multi'
                                    ? 'bg-amber-950/10 dark:bg-slate-900 border-amber-500/40 hover:border-amber-400'
                                    : item.status === 'single'
                                    ? 'bg-emerald-950/10 dark:bg-slate-900 border-emerald-500/40 hover:border-emerald-400'
                                    : 'bg-rose-950/10 dark:bg-slate-900 border-rose-500/40 hover:border-rose-400'
                                }`}
                              >
                                {/* Card Header Box */}
                                <div className="flex items-center justify-between pb-2.5 border-b border-slate-200 dark:border-slate-800">
                                  <div className="flex items-center gap-2">
                                    <input
                                      type="checkbox"
                                      className="rounded border-slate-300 text-indigo-600 focus:ring-0 cursor-pointer"
                                      checked={isSelected}
                                      onChange={() => handleToggleSelectRow(item.machine)}
                                    />
                                    <div>
                                      <span className="font-mono font-black text-slate-900 dark:text-white text-base block leading-none">
                                        {item.machine}
                                      </span>
                                      <span className="text-[10px] text-slate-400 font-mono mt-1 block">
                                        بطاقة ماكينة #{globalIdx}
                                      </span>
                                    </div>
                                  </div>
                                  <span
                                    className={`px-2 py-1 rounded-xl text-[10px] font-black border ${
                                      item.status === 'single'
                                        ? 'bg-emerald-500/20 text-emerald-400 border-emerald-500/30'
                                        : item.status === 'multi'
                                        ? 'bg-amber-500/20 text-amber-400 border-amber-500/30'
                                        : 'bg-rose-500/20 text-rose-400 border-rose-500/30'
                                    }`}
                                  >
                                    {item.status === 'single'
                                      ? 'مندوب واحد (1/1)'
                                      : item.status === 'multi'
                                      ? `مشتركة (${item.repOrder}/${item.totalRepsForMachine})`
                                      : 'شاغرة'}
                                  </span>
                                </div>

                                {/* Rep Sub-Card Box */}
                                <div className="p-3 rounded-2xl bg-slate-50 dark:bg-slate-800/60 border border-slate-200 dark:border-slate-700/60 space-y-2">
                                  <div className="flex items-center justify-between text-xs">
                                    <span className="text-slate-500 dark:text-slate-400 font-bold text-[10px]">
                                      بيانات المندوب:
                                    </span>
                                    {item.account !== 'غير متوفر' ? (
                                      <button
                                        onClick={() => {
                                          setRepLookupQuery(item.account);
                                          setIsRepLookupOpen(true);
                                        }}
                                        className="font-mono text-xs font-black px-2 py-0.5 rounded-md bg-indigo-600/20 text-indigo-400 border border-indigo-500/30 hover:bg-indigo-600 hover:text-white transition-all cursor-pointer"
                                        title="اضغط لاستعلام كافة ماكينات هذا المندوب"
                                      >
                                        حساب: {item.account}
                                      </button>
                                    ) : (
                                      <span className="text-[10px] text-slate-400 font-mono">بدون كود</span>
                                    )}
                                  </div>

                                  <div className="flex items-center gap-2">
                                    <div className="w-8 h-8 rounded-xl bg-indigo-500/20 text-indigo-400 flex items-center justify-center font-bold text-xs shrink-0">
                                      <UserCheck className="w-4 h-4" />
                                    </div>
                                    <div className="min-w-0 flex-1">
                                      <p className="font-black text-xs sm:text-sm text-slate-900 dark:text-white truncate">
                                        {item.repName || 'لا يوجد مندوب'}
                                      </p>
                                      {item.isMissingRepName && (
                                        <span className="text-[10px] text-amber-500 font-bold block">
                                          ⚠️ الاسم غير مدون بالشيت
                                        </span>
                                      )}
                                    </div>
                                  </div>
                                </div>

                                {/* Card Footer Actions */}
                                <div className="pt-2 flex items-center justify-between text-[11px] border-t border-slate-100 dark:border-slate-800">
                                  <span className="text-slate-400 text-[10px]">
                                    {item.status === 'multi'
                                      ? `مندوب رقم ${item.repOrder} من ${item.totalRepsForMachine}`
                                      : 'ماكينة فردية'}
                                  </span>
                                  <button
                                    onClick={() => {
                                      setRepLookupQuery(
                                        item.account !== 'غير متوفر' ? item.account : item.repName
                                      );
                                      setIsRepLookupOpen(true);
                                    }}
                                    className="text-xs font-bold text-indigo-500 hover:text-indigo-400 flex items-center gap-1 cursor-pointer"
                                  >
                                    <span>كشف المندوب</span>
                                    <ChevronLeft className="w-3.5 h-3.5" />
                                  </button>
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    </div>
                  )}
                </div>
              ) : (
                /* Table View */
                <div
                  ref={tableContainerRef}
                  onScroll={(e) => setTableScrollTop(e.currentTarget.scrollTop)}
                  className="max-h-[550px] overflow-y-auto border border-slate-800/90 rounded-2xl bg-slate-900/80 backdrop-blur-md shadow-2xl relative"
                >
                  <table className="w-full text-right border-collapse text-xs">
                    <thead className="bg-[#0a0f1d] text-slate-400 font-bold border-b border-slate-800 sticky top-0 z-20 shadow-md">
                      <tr>
                        <th className="p-3.5 w-12 text-center">
                          <input
                            type="checkbox"
                            className="w-4 h-4 rounded border-slate-700 bg-slate-800 text-indigo-600 focus:ring-0 cursor-pointer"
                            checked={
                              displayedSlice.length > 0 &&
                              displayedSlice.every((r) => selectedMachines.has(r.machine))
                            }
                            onChange={(e) => handleToggleSelectAll(e.target.checked)}
                            title="تحديد الكل المعروض"
                          />
                        </th>
                        <th className="p-3.5 w-14 text-center font-mono">م</th>
                        <th className="p-3.5 w-48">رقم الماكينة</th>
                        <th className="p-3.5 w-40">رقم حساب المندوب</th>
                        <th className="p-3.5">اسم المندوب المسند</th>
                        <th className="p-3.5 w-48 text-center">ترتيب المندوب للماكينة</th>
                        <th className="p-3.5 w-40 text-center">حالة الماكينة</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-800/80 text-slate-300">
                      {displayedSlice.length === 0 ? (
                        <tr>
                          <td colSpan={7} className="p-8 text-center text-slate-400">
                            لا توجد نتائج تطابق شرط البحث أو الفلتر. ارفع الشيتات أو اختر «عينة فورية».
                          </td>
                        </tr>
                      ) : (
                        <>
                          {tableVirtualState.topPadding > 0 && (
                            <tr style={{ border: 'none' }}>
                              <td colSpan={7} style={{ height: tableVirtualState.topPadding, padding: 0, border: 'none' }} />
                            </tr>
                          )}
                          {displayedSlice.slice(tableVirtualState.startIndex, tableVirtualState.endIndex).map((item, idx) => {
                            const globalIdx = (currentPage - 1) * pageSize + tableVirtualState.startIndex + idx + 1;
                            const isSelected = selectedMachines.has(item.machine);
                            const isChildRow = item.status === 'multi' && item.repOrder > 1;

                            const rowClass = isChildRow
                              ? 'bg-slate-950/60 hover:bg-slate-800/40 transition-colors'
                              : 'hover:bg-slate-800/50 transition-colors';

                            return (
                              <tr key={item.id} className={rowClass} style={{ height: tableRowHeight }}>
                                <td className="p-3.5 text-center">
                                  <input
                                    type="checkbox"
                                    className="w-4 h-4 rounded border-slate-700 bg-slate-800 text-indigo-600 cursor-pointer"
                                    checked={isSelected}
                                    onChange={() => handleToggleSelectRow(item.machine)}
                                  />
                                </td>
                                <td className="p-3.5 text-center text-slate-500 font-mono font-bold">
                                  {globalIdx}
                                </td>

                                {/* Machine Column */}
                                <td className="p-3.5">
                                  {isChildRow ? (
                                    <div className="flex items-center gap-2 font-mono font-bold text-slate-300 pr-3 border-r-2 border-indigo-500/50">
                                      <ArrowDown className="w-3.5 h-3.5 text-indigo-400" />
                                      <span className="tracking-wide">{item.machine}</span>
                                      <span className="text-[10px] text-slate-400 font-sans font-normal px-1.5 py-0.5 rounded bg-slate-800/80 border border-slate-700/50">
                                        (تابع)
                                      </span>
                                    </div>
                                  ) : (
                                    <span className="font-mono font-black text-white text-sm tracking-wide">
                                      {item.machine}
                                    </span>
                                  )}
                                </td>

                                {/* Account Column */}
                                <td className="p-3.5">
                                  {item.account === 'غير متوفر' ? (
                                    <span className="text-slate-500 font-mono text-xs">غير متوفر</span>
                                  ) : (
                                    <button
                                      onClick={() => {
                                        setRepLookupQuery(item.account);
                                        setIsRepLookupOpen(true);
                                      }}
                                      className="font-mono text-xs font-bold px-3 py-1 rounded-lg bg-blue-950/60 text-blue-400 border border-blue-500/30 hover:bg-blue-900/60 transition-colors cursor-pointer"
                                      title="اضغط لمعاينة كافة ماكينات هذا المندوب وتصدير كشف خاص به"
                                    >
                                      {item.account}
                                    </button>
                                  )}
                                </td>

                                {/* Rep Name Column */}
                                <td className="p-3.5">
                                  {item.status === 'none' || item.isMissingRepName ? (
                                    <div className="flex items-center gap-2.5">
                                      <div className="w-7 h-7 rounded-full bg-slate-800 border border-slate-700/60 text-rose-400 flex items-center justify-center text-xs shadow-inner">
                                        <UserX className="w-3.5 h-3.5" />
                                      </div>
                                      <span className="text-xs font-bold text-rose-400">
                                        {item.repName || 'لا يوجد مندوب مسند'}
                                      </span>
                                    </div>
                                  ) : (
                                    <button
                                      onClick={() => {
                                        setRepLookupQuery(item.account !== 'غير متوفر' ? item.account : item.repName);
                                        setIsRepLookupOpen(true);
                                      }}
                                      className="flex items-center gap-2.5 hover:opacity-80 transition-opacity cursor-pointer group text-right"
                                      title="اضغط لمعاينة كافة ماكينات هذا المندوب"
                                    >
                                      <div className="w-7 h-7 rounded-full bg-slate-800 border border-slate-700/60 text-blue-400 flex items-center justify-center text-xs shadow-inner group-hover:bg-blue-600 group-hover:text-white transition-colors">
                                        <UserCheck className="w-3.5 h-3.5" />
                                      </div>
                                      <span className="font-bold text-white text-xs group-hover:text-blue-300 transition-colors">
                                        {item.repName}
                                      </span>
                                    </button>
                                  )}
                                </td>

                                {/* Rep Order Column */}
                                <td className="p-3.5 text-center">
                                  {item.status === 'none' ? (
                                    <span className="text-xs text-slate-500 font-bold">-</span>
                                  ) : item.status === 'single' ? (
                                    <span className="inline-flex items-center px-3 py-1 rounded-full text-[11px] font-bold bg-emerald-500/10 text-emerald-400 border border-emerald-500/20">
                                      مندوب رئيسي (1 من 1)
                                    </span>
                                  ) : item.repOrder === 1 ? (
                                    <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-[11px] font-bold bg-amber-500/10 text-amber-400 border border-amber-500/20">
                                      المندوب الأول (1 من {item.totalRepsForMachine})
                                    </span>
                                  ) : (
                                    <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-[11px] font-bold bg-slate-800 text-slate-300 border border-slate-700/60">
                                      مندوب إضافي ({item.repOrder} من {item.totalRepsForMachine})
                                    </span>
                                  )}
                                </td>

                                {/* Machine Status Column */}
                                <td className="p-3.5 text-center">
                                  {item.status === 'single' ? (
                                    <span className="inline-flex items-center px-3 py-1 rounded-lg text-[11px] font-bold bg-emerald-500/10 text-emerald-400 border border-emerald-500/20">
                                      مندوب فردي
                                    </span>
                                  ) : item.status === 'multi' ? (
                                    <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-lg text-[11px] font-bold bg-indigo-500/10 text-indigo-400 border border-indigo-500/20">
                                      <CircleAlert className="w-3.5 h-3.5 text-indigo-400" />
                                      <span>متعددة ({item.totalRepsForMachine} مناديب)</span>
                                    </span>
                                  ) : (
                                    <span className="inline-flex items-center gap-1 px-3 py-1 rounded-lg text-[11px] font-bold bg-rose-500/10 text-rose-400 border border-rose-500/20">
                                      شاغرة
                                    </span>
                                  )}
                                </td>
                              </tr>
                            );
                          })}
                          {tableVirtualState.bottomPadding > 0 && (
                            <tr style={{ border: 'none' }}>
                              <td colSpan={7} style={{ height: tableVirtualState.bottomPadding, padding: 0, border: 'none' }} />
                            </tr>
                          )}
                        </>
                      )}
                    </tbody>
                  </table>
                </div>
              )}

              {/* Pagination */}
              <div className="flex items-center justify-between text-xs text-slate-500 pt-2">
                <div>
                  عرض{' '}
                  <span className="font-bold text-slate-800 dark:text-white">
                    {displayedSlice.length}
                  </span>{' '}
                  من{' '}
                  <span className="font-bold text-slate-800 dark:text-white">
                    {filteredRows.length}
                  </span>{' '}
                  سطر مفصل
                </div>
                <div className="flex items-center gap-1 font-bold">
                  <button
                    disabled={currentPage <= 1}
                    onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
                    className="p-1 px-2.5 rounded bg-slate-100 dark:bg-slate-800 disabled:opacity-40 flex items-center gap-1 cursor-pointer"
                  >
                    <ChevronRight className="w-3.5 h-3.5" />
                    <span>السابق</span>
                  </button>
                  <span className="px-2">
                    صفحة {currentPage} من {maxPages}
                  </span>
                  <button
                    disabled={currentPage >= maxPages}
                    onClick={() => setCurrentPage((p) => Math.min(maxPages, p + 1))}
                    className="p-1 px-2.5 rounded bg-slate-100 dark:bg-slate-800 disabled:opacity-40 flex items-center gap-1 cursor-pointer"
                  >
                    <span>التالي</span>
                    <ChevronLeft className="w-3.5 h-3.5" />
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* TAB 5: Audit & Conflict Inspection */}
        {currentTab === 5 && (
          <AuditConflictsTab
            duplicates={duplicates}
            irregulars={irregulars}
            onFilterDuplicates={() => {
              if (duplicates.length > 0) {
                setSearchQuery(duplicates[0].machine);
                setCurrentFilter('all');
                setCurrentTab(4);
                showToast(`🔍 تم تصفية الماكينة المكررة: ${duplicates[0].machine}`);
              }
            }}
            onFilterIrregulars={() => {
              if (irregulars.length > 0) {
                setSearchQuery(irregulars[0].account);
                setCurrentFilter('all');
                setCurrentTab(4);
                showToast(`🔍 تم تصفية الحساب الشاذ: ${irregulars[0].account}`);
              }
            }}
          />
        )}

        {/* TAB 6: Standalone Desktop App & HTML Exporter */}
        {currentTab === 6 && (
          <DesktopAppTab
            sheet1={sheet1}
            sheet2={sheet2}
            expandedRows={expandedRows}
            machinesResults={machinesResults}
            onOpenRepLookup={() => setIsRepLookupOpen(true)}
            onOpenManualEditor={(mode) => {
              setEditorInitialMode(mode);
              setIsManualEditorOpen(true);
            }}
          />
        )}
      </main>

      {/* Bulk Floating Action Bar */}
      <BulkActionBar
        selectedCount={selectedMachines.size}
        onClearSelection={() => setSelectedMachines(new Set())}
        onReassign={handleBulkReassign}
        onMarkVacant={handleBulkMarkVacant}
      />
    </div>
  );
}
