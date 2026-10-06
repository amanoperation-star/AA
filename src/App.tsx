import React, { useState, useEffect, useMemo, useRef } from 'react';
import { motion, AnimatePresence } from 'motion/react';
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
  Eye,
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

function normalizeMachineId(id: string | number | undefined | null): string {
  if (id === undefined || id === null) return '';
  let clean = String(id).trim();
  // Remove "7-" or "٧-" prefixes
  if (clean.startsWith('7-')) {
    clean = clean.substring(2).trim();
  } else if (clean.startsWith('٧-')) {
    clean = clean.substring(2).trim();
  }
  return clean;
}

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

  // Sheets data (1: Machines, 2: Payments, 3: Cash)
  const [sheet1, setSheet1] = useState<SheetRow[]>([]);
  const [sheet2, setSheet2] = useState<SheetRow[]>([]);
  const [sheet3, setSheet3] = useState<SheetRow[]>([]);
  const [sheet1FileName, setSheet1FileName] = useState<string>('');
  const [sheet2FileName, setSheet2FileName] = useState<string>('');
  const [sheet3FileName, setSheet3FileName] = useState<string>('');

  // Upload card interactive states & refs
  const [isSheet1Loading, setIsSheet1Loading] = useState<boolean>(false);
  const [isSheet2Loading, setIsSheet2Loading] = useState<boolean>(false);
  const [isSheet3Loading, setIsSheet3Loading] = useState<boolean>(false);
  const [isDragging1, setIsDragging1] = useState<boolean>(false);
  const [isDragging2, setIsDragging2] = useState<boolean>(false);
  const [isDragging3, setIsDragging3] = useState<boolean>(false);
  const sheet1InputRef = useRef<HTMLInputElement>(null);
  const sheet2InputRef = useRef<HTMLInputElement>(null);
  const sheet3InputRef = useRef<HTMLInputElement>(null);

  // Column mapping (Machines, Payments, Cash)
  const [colM1, setColM1] = useState<string>('');
  const [colM2, setColM2] = useState<string>('');
  const [colAcc, setColAcc] = useState<string>('');
  const [colRep, setColRep] = useState<string>('');
  const [colM3, setColM3] = useState<string>('');
  const [colCashAcc, setColCashAcc] = useState<string>('');
  const [colCashRep, setColCashRep] = useState<string>('');
  const [emptyRepFallback, setEmptyRepFallback] = useState<string>('لا يوجد مسؤول');

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
  const [sheet1PreviewFilter, setSheet1PreviewFilter] = useState<'all' | 'cash' | 'payments'>('all');

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

    if (sheet3.length > 0) {
      const cols3 = Object.keys(sheet3[0]);
      const foundM3 = cols3.find((c) => /ماكينة|جهاز|pos|machine|sn/i.test(c)) || cols3[0];
      const foundCashAcc = cols3.find((c) => /حساب|كود|رقم.*مندوب|account|acc|code/i.test(c)) || '';
      let foundCashRep = cols3.find((c) => /اسم.*مندوب|اسم.*العميل|اسم.*المستخدم|اسم|rep.*name|agent.*name|name/i.test(c));
      if (!foundCashRep) {
        foundCashRep = cols3.find((c) => /(مندوب|agent|rep)/i.test(c) && !/(رقم|حساب|كود|acc|code|id)/i.test(c));
      }
      if (!foundCashRep) {
        foundCashRep = cols3.find((c) => /(مندوب|agent|rep)/i.test(c)) || cols3[1] || cols3[0];
      }

      setColM3(foundM3);
      setColCashAcc(foundCashAcc);
      setColCashRep(foundCashRep);
    }
  }, [sheet1, sheet2, sheet3]);

  // Core Reconciliation Engine (Supports Sheet 1: Machines, Sheet 2: Payments, Sheet 3: Cash)
  const executeReconciliation = (
    s1: SheetRow[],
    s2: SheetRow[],
    s3: SheetRow[],
    m1Name: string,
    m2Name: string,
    accName: string,
    repNameCol: string,
    m3Name: string,
    cashAccName: string,
    cashRepNameCol: string,
    fallbackVal = emptyRepFallback
  ) => {
    const repMap = new Map<string, Array<{ account: string; name: string; isMissingRepName: boolean; type: 'payment' | 'cash' }>>();
    const irregularsList: IrregularAccountItem[] = [];

    // Process Sheet 2 (Payments)
    for (let i = 0; i < s2.length; i++) {
      const row = s2[i];
      const rawM = row[m2Name];
      if (rawM !== undefined && rawM !== null && rawM !== '') {
        const rawMStr = String(rawM).trim();
        const key = normalizeMachineId(rawMStr);
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

        if (repAcc !== 'غير متوفر') {
          const hasSpecialChars = /[^0-9\-_]/i.test(repAcc);
          const hasSpaces = /\s/.test(repAcc);
          const isUnusuallyShort = repAcc.length < 2;
          if (hasSpecialChars || hasSpaces || isUnusuallyShort) {
            irregularsList.push({
              machine: rawMStr,
              account: repAcc,
              repName,
              reason: hasSpaces ? 'يحتوي على مسافات (مدفوعات)' : hasSpecialChars ? 'يحتوي على رموز/حروف (مدفوعات)' : 'قصير جداً',
            });
          }
        }

        if (!repMap.has(key)) repMap.set(key, []);
        repMap.get(key)!.push({ account: repAcc, name: repName, isMissingRepName: isRepNameEmpty, type: 'payment' });
      }
    }

    // Process Sheet 3 (Cash)
    for (let i = 0; i < s3.length; i++) {
      const row = s3[i];
      const rawM = row[m3Name];
      if (rawM !== undefined && rawM !== null && rawM !== '') {
        const rawMStr = String(rawM).trim();
        const key = normalizeMachineId(rawMStr);
        const rawRepName = row[cashRepNameCol];
        const isRepNameEmpty =
          rawRepName === undefined ||
          rawRepName === null ||
          String(rawRepName).trim() === '' ||
          String(rawRepName).trim() === '-' ||
          String(rawRepName).trim() === 'null';

        const repName = isRepNameEmpty ? fallbackVal : String(rawRepName).trim();
        const repAcc =
          cashAccName && row[cashAccName] !== undefined && row[cashAccName] !== ''
            ? String(row[cashAccName]).trim()
            : 'غير متوفر';

        if (repAcc !== 'غير متوفر') {
          const hasSpecialChars = /[^0-9\-_]/i.test(repAcc);
          const hasSpaces = /\s/.test(repAcc);
          const isUnusuallyShort = repAcc.length < 2;
          if (hasSpecialChars || hasSpaces || isUnusuallyShort) {
            irregularsList.push({
              machine: rawMStr,
              account: repAcc,
              repName,
              reason: hasSpaces ? 'يحتوي على مسافات (كاش)' : hasSpecialChars ? 'يحتوي على رموز/حروف (كاش)' : 'قصير جداً',
            });
          }
        }

        if (!repMap.has(key)) repMap.set(key, []);
        repMap.get(key)!.push({ account: repAcc, name: repName, isMissingRepName: isRepNameEmpty, type: 'cash' });
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

      const originalMachineId = String(mVal).trim();
      const lookupKey = normalizeMachineId(originalMachineId);
      const reps = repMap.get(lookupKey) || [];
      const count = reps.length;

      let status: 'single' | 'multi' | 'none' = 'none';
      if (count === 1) status = 'single';
      else if (count > 1) status = 'multi';

      // Check if machine starts with 7- or matches a cash representative
      const isCashMachine = originalMachineId.startsWith('7-') || originalMachineId.startsWith('٧-') || reps.some(r => r.type === 'cash');
      const cleanMachineId = normalizeMachineId(originalMachineId);

      summaries.push({
        index: i + 1,
        machine: originalMachineId,
        repCount: count,
        reps,
        status,
      });

      if (count === 0) {
        rowCounter++;
        const isVacantCash = row._tabType === 'cash' || isCashMachine;
        const vacantMachineDisplay = isVacantCash
          ? (originalMachineId.startsWith('7-') || originalMachineId.startsWith('٧-') ? originalMachineId : `7-${cleanMachineId}`)
          : cleanMachineId;

        flatRows.push({
          id: `row-${rowCounter}`,
          machine: vacantMachineDisplay,
          account: 'غير متوفر',
          repName: fallbackVal,
          repOrder: 0,
          totalRepsForMachine: 0,
          status: 'none',
          machineGroupIndex: i + 1,
          isFirstOfGroup: true,
          isLastOfGroup: true,
          isMissingRepName: true,
          type: isVacantCash ? 'cash' : 'payment',
        });
      } else {
        for (let rIdx = 0; rIdx < count; rIdx++) {
          rowCounter++;
          const r = reps[rIdx];

          // CRITICAL ACCURACY FIX:
          // Payment accounts (r.type === 'payment') MUST display clean machine ID without 7- prefix
          // Cash accounts (r.type === 'cash') MUST display machine ID with 7- prefix
          const repMachineId = r.type === 'payment'
            ? cleanMachineId
            : (cleanMachineId.startsWith('7-') || cleanMachineId.startsWith('٧-') ? cleanMachineId : `7-${cleanMachineId}`);

          flatRows.push({
            id: `row-${rowCounter}`,
            machine: repMachineId,
            account: r.account,
            repName: r.name,
            repOrder: rIdx + 1,
            totalRepsForMachine: count,
            status,
            machineGroupIndex: i + 1,
            isFirstOfGroup: rIdx === 0,
            isLastOfGroup: rIdx === count - 1,
            isMissingRepName: r.isMissingRepName,
            type: r.type,
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
    if (!sheet1.length) {
      showToast('⚠️ يرجى رفع شيت الماكينات أولاً، أو اضغط على «عينة فورية».');
      return;
    }

    if (!sheet2.length && !sheet3.length) {
      showToast('⚠️ يرجى رفع شيت المدفوعات أو شيت الكاش للمطابقة مع شيت الماكينات.');
      return;
    }

    if (!colM1) {
      showToast('⚠️ يرجى ضبط عمود الماكينة لشيت الماكينات في تبويب ضبط الأعمدة.');
      setCurrentTab(4);
      return;
    }

    if (sheet2.length > 0 && (!colM2 || !colRep)) {
      showToast('⚠️ يرجى ضبط أعمدة الماكينة والمندوب لشيت المدفوعات في تبويب ضبط الأعمدة.');
      setCurrentTab(4);
      return;
    }

    if (sheet3.length > 0 && (!colM3 || !colCashRep)) {
      showToast('⚠️ يرجى ضبط أعمدة الماكينة ومندوب الكاش لشيت الكاش في تبويب ضبط الأعمدة.');
      setCurrentTab(4);
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
        executeReconciliation(sheet1, sheet2, sheet3, colM1, colM2, colAcc, colRep, colM3, colCashAcc, colCashRep, emptyRepFallback);
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
    setIsSheet1Loading(true);
    setProgressState({
      isOpen: true,
      title: 'قراءة',
      fileName: file.name,
      subtitle: 'جاري تحليل خلايا وسجلات ملف الإكسل...',
      percent: 40,
      stepText: 'استخراج أرقام الماكينات وتدقيق التكرارات...',
      countText: '',
      isComplete: false,
    });

    try {
      const data = await parseExcelFile(file);
      if (!data || data.length === 0) {
        setIsSheet1Loading(false);
        setProgressState((p) => ({ ...p, isOpen: false }));
        showToast('⚠️ لم يتم العثور على أسطر بيانات صالحة في «' + file.name + '». يرجى التأكد من محتوى الملف.');
        return;
      }

      setSheet1(data);
      setSheet1FileName(file.name);

      setProgressState((p) => ({
        ...p,
        percent: 100,
        stepText: `تم قراءة ${data.length} ماكينة بنجاح!`,
        isComplete: true,
      }));

      setTimeout(() => {
        setIsSheet1Loading(false);
        setProgressState((p) => ({ ...p, isOpen: false }));
        showToast(`✅ تم قراءة شيت الماكينات بنجاح (${data.length.toLocaleString('ar-EG')} ماكينة).`);
        const canAutoRun = (sheet2.length > 0 || sheet3.length > 0) &&
                           colM1 &&
                           (sheet2.length === 0 || (colM2 && colRep)) &&
                           (sheet3.length === 0 || (colM3 && colCashRep));
        if (canAutoRun) {
          executeReconciliation(data, sheet2, sheet3, colM1, colM2, colAcc, colRep, colM3, colCashAcc, colCashRep, emptyRepFallback);
        }
      }, 300);
    } catch (err: any) {
      console.error(err);
      setIsSheet1Loading(false);
      setProgressState((p) => ({ ...p, isOpen: false }));
      showToast(`❌ تعذر قراءة الملف: ${err?.message || 'يرجى التأكد من صيغة Excel أو CSV.'}`);
    }
  };

  // Handle Sheet 2 upload (Payments Sheet)
  const handleSheet2Upload = async (file: File) => {
    setIsSheet2Loading(true);
    setProgressState({
      isOpen: true,
      title: 'قراءة',
      fileName: file.name,
      subtitle: 'جاري تحليل خلايا وسجلات ملف الإكسل...',
      percent: 40,
      stepText: 'استخراج حسابات المدفوعات وربطها بالماكينات...',
      countText: '',
      isComplete: false,
    });

    try {
      const data = await parseExcelFile(file);
      if (!data || data.length === 0) {
        setIsSheet2Loading(false);
        setProgressState((p) => ({ ...p, isOpen: false }));
        showToast('⚠️ لم يتم العثور على أسطر بيانات صالحة في شيت المدفوعات «' + file.name + '».');
        return;
      }

      setSheet2(data);
      setSheet2FileName(file.name);

      setProgressState((p) => ({
        ...p,
        percent: 100,
        stepText: `تم قراءة ${data.length} سجل مدفوعات بنجاح!`,
        isComplete: true,
      }));

      setTimeout(() => {
        setIsSheet2Loading(false);
        setProgressState((p) => ({ ...p, isOpen: false }));
        showToast(`✅ تم قراءة شيت المدفوعات بنجاح (${data.length.toLocaleString('ar-EG')} سجل).`);
        const canAutoRun = sheet1.length > 0 &&
                           colM1 &&
                           (colM2 && colRep) &&
                           (sheet3.length === 0 || (colM3 && colCashRep));
        if (canAutoRun) {
          executeReconciliation(sheet1, data, sheet3, colM1, colM2, colAcc, colRep, colM3, colCashAcc, colCashRep, emptyRepFallback);
        }
      }, 300);
    } catch (err: any) {
      console.error(err);
      setIsSheet2Loading(false);
      setProgressState((p) => ({ ...p, isOpen: false }));
      showToast(`❌ تعذر قراءة شيت المدفوعات: ${err?.message || 'يرجى التأكد من صيغة الملف.'}`);
    }
  };

  // Handle Sheet 3 upload (Cash Sheet)
  const handleSheet3Upload = async (file: File) => {
    setIsSheet3Loading(true);
    setProgressState({
      isOpen: true,
      title: 'قراءة',
      fileName: file.name,
      subtitle: 'جاري تحليل خلايا وسجلات شيت الكاش...',
      percent: 40,
      stepText: 'استخراج حسابات الكاش وربطها بالماكينات...',
      countText: '',
      isComplete: false,
    });

    try {
      const data = await parseExcelFile(file);
      if (!data || data.length === 0) {
        setIsSheet3Loading(false);
        setProgressState((p) => ({ ...p, isOpen: false }));
        showToast('⚠️ لم يتم العثور على أسطر بيانات صالحة في شيت الكاش «' + file.name + '».');
        return;
      }

      setSheet3(data);
      setSheet3FileName(file.name);

      setProgressState((p) => ({
        ...p,
        percent: 100,
        stepText: `تم قراءة ${data.length} سجل كاش بنجاح!`,
        isComplete: true,
      }));

      setTimeout(() => {
        setIsSheet3Loading(false);
        setProgressState((p) => ({ ...p, isOpen: false }));
        showToast(`✅ تم قراءة شيت الكاش بنجاح (${data.length.toLocaleString('ar-EG')} سجل).`);
        const canAutoRun = sheet1.length > 0 &&
                           colM1 &&
                           (sheet2.length === 0 || (colM2 && colRep)) &&
                           (colM3 && colCashRep);
        if (canAutoRun) {
          executeReconciliation(sheet1, sheet2, data, colM1, colM2, colAcc, colRep, colM3, colCashAcc, colCashRep, emptyRepFallback);
        }
      }, 300);
    } catch (err: any) {
      console.error(err);
      setIsSheet3Loading(false);
      setProgressState((p) => ({ ...p, isOpen: false }));
      showToast(`❌ تعذر قراءة شيت الكاش: ${err?.message || 'يرجى التأكد من صيغة الملف.'}`);
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
        { 'رقم الماكينة': 'POS-100', 'رقم حساب المدفوعات': '1234', 'اسم مسؤول المدفوعات': 'أحمد محمود سالم' },
        { 'رقم الماكينة': 'POS-100', 'رقم حساب المدفوعات': '456', 'اسم مسؤول المدفوعات': 'محمود حسن رضوان' },
        { 'رقم الماكينة': 'POS-200', 'رقم حساب المدفوعات': '555', 'اسم مسؤول المدفوعات': '' },
        { 'رقم الماكينة': 'POS-300', 'رقم حساب المدفوعات': '789', 'اسم مسؤول المدفوعات': 'خالد عبد الرحمن' },
        { 'رقم الماكينة': 'POS-300', 'رقم حساب المدفوعات': '890', 'اسم مسؤول المدفوعات': 'طارق زياد العتيبي' },
        { 'رقم الماكينة': 'POS-300', 'رقم حساب المدفوعات': '999', 'اسم مسؤول المدفوعات': 'عمر فاروق الشامي' },
        { 'رقم الماكينة': 'POS-400', 'رقم حساب المدفوعات': '777-ERR', 'اسم مسؤول المدفوعات': 'إبراهيم حسني مراد' },
      ];

      const sample3: SheetRow[] = [
        { 'رقم الماكينة': '7-POS-100', 'رقم حساب الكاش': '1234', 'اسم مسؤول الكاش': 'أحمد محمود سالم (كاش)' },
        { 'رقم الماكينة': '7-POS-200', 'رقم حساب الكاش': '555', 'اسم مسؤول الكاش': 'خالد عبد الرحمن (كاش)' },
        { 'رقم الماكينة': '7-POS-300', 'رقم حساب الكاش': '789', 'اسم مسؤول الكاش': 'طارق زياد العتيبي (كاش)' },
        { 'رقم الماكينة': '7-POS-400', 'رقم حساب الكاش': '777', 'اسم مسؤول الكاش': 'سامح عبد الله كمال (كاش)' },
        { 'رقم الماكينة': '7-POS-500', 'رقم حساب الكاش': '888', 'اسم مسؤول الكاش': 'مصطفى كمال (كاش)' },
      ];

      setSheet1(sample1);
      setSheet1FileName('عينة_شيت_الماكينات.xlsx');
      setSheet2(sample2);
      setSheet2FileName('عينة_شيت_المدفوعات.xlsx');
      setSheet3(sample3);
      setSheet3FileName('عينة_شيت_الكاش.xlsx');

      const m1 = 'رقم الماكينة';
      const m2 = 'رقم الماكينة';
      const acc = 'رقم حساب المدفوعات';
      const rep = 'اسم مسؤول المدفوعات';
      const m3 = 'رقم الماكينة';
      const cashAcc = 'رقم حساب الكاش';
      const cashRep = 'اسم مسؤول الكاش';

      setColM1(m1);
      setColM2(m2);
      setColAcc(acc);
      setColRep(rep);
      setColM3(m3);
      setColCashAcc(cashAcc);
      setColCashRep(cashRep);

      executeReconciliation(sample1, sample2, sample3, m1, m2, acc, rep, m3, cashAcc, cashRep, 'لا يوجد مسؤول');

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
    setSheet3([]);
    setSheet1FileName('');
    setSheet2FileName('');
    setSheet3FileName('');
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

  // Automatically prepend "7-" to all machine numbers in Cash sheet (Sheet 3) and/or Machines sheet (Sheet 1)
  const handleApplyCashPrefix = () => {
    if (!sheet1.length && !sheet3.length) {
      showToast('⚠️ يرجى رفع شيت الماكينات أو شيت الكاش أولاً لتطبيق البادئة.');
      return;
    }

    let modifiedCount = 0;

    // 1. Update Sheet 3 (Cash Sheet)
    let updatedSheet3 = sheet3;
    if (sheet3.length > 0) {
      const mCol3 = colM3 || Object.keys(sheet3[0])[0];
      if (mCol3) {
        updatedSheet3 = sheet3.map((row) => {
          const val = row[mCol3];
          if (val !== undefined && val !== null && val !== '') {
            const str = String(val).trim();
            if (!str.startsWith('7-') && !str.startsWith('٧-')) {
              modifiedCount++;
              return {
                ...row,
                [mCol3]: `7-${str}`,
              };
            }
          }
          return row;
        });
        setSheet3(updatedSheet3);
      }
    }

    // 2. Update Sheet 1 (Machines Sheet)
    let updatedSheet1 = sheet1;
    if (sheet1.length > 0) {
      const mCol1 = colM1 || Object.keys(sheet1[0])[0];
      if (mCol1) {
        const cashMachines = updatedSheet3.length > 0
          ? new Set(updatedSheet3.map((r) => normalizeMachineId(String(r[colM3 || Object.keys(r)[0]]))))
          : null;

        updatedSheet1 = sheet1.map((row) => {
          const val = row[mCol1];
          if (val !== undefined && val !== null && val !== '') {
            const str = String(val).trim();
            if (!str.startsWith('7-') && !str.startsWith('٧-')) {
              // Convert ONLY cash machines; NEVER convert payment machines
              const isExplicitCash = row._tabType === 'cash';
              const isExplicitPayment = row._tabType === 'payments';
              const matchesSheet3 = cashMachines ? cashMachines.has(normalizeMachineId(str)) : false;

              if (isExplicitCash || matchesSheet3 || (!cashMachines && !isExplicitPayment && row._tabType !== 'payments')) {
                modifiedCount++;
                return {
                  ...row,
                  [mCol1]: `7-${str}`,
                };
              }
            }
          }
          return row;
        });
        setSheet1(updatedSheet1);
      }
    }

    if (modifiedCount > 0) {
      showToast(`✅ تم تحويل الماكينات وإضافة البادئة 7- بنجاح! (تم تعديل ${modifiedCount} ماكينة)`);
    } else {
      showToast(`✅ كافة ماكينات الكاش تحتوي بالفعل على البادئة 7-`);
    }

    // Auto run reconciliation if sheets are ready
    if (updatedSheet1.length > 0 && (colM1 || Object.keys(updatedSheet1[0])[0])) {
      const activeColM1 = colM1 || Object.keys(updatedSheet1[0])[0];
      executeReconciliation(
        updatedSheet1,
        sheet2,
        updatedSheet3,
        activeColM1,
        colM2,
        colAcc,
        colRep,
        colM3,
        colCashAcc,
        colCashRep,
        emptyRepFallback
      );
    }
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
    executeReconciliation(sheet1, updatedSheet2, sheet3, colM1, colM2, colAcc, colRep, colM3, colCashAcc, colCashRep, emptyRepFallback);
    setSelectedMachines(new Set());
    showToast(`✅ تم إعادة توزيع ${selectedMachines.size} ماكينة للمندوب بنجاح!`);
  };

  const handleBulkMarkVacant = () => {
    if (selectedMachines.size === 0) return;
    const updatedSheet2 = sheet2.filter((r) => !selectedMachines.has(String(r[colM2]).trim()));
    setSheet2(updatedSheet2);
    executeReconciliation(sheet1, updatedSheet2, sheet3, colM1, colM2, colAcc, colRep, colM3, colCashAcc, colCashRep, emptyRepFallback);
    setSelectedMachines(new Set());
    showToast(`✅ تم تحويل ${selectedMachines.size} ماكينة إلى شاغرة (بدون مندوب).`);
  };

  // Sheet 1 Preview filters & counts
  const cashMachinesCount = useMemo(() => {
    return sheet1.filter((r) => {
      const mCol = colM1 || Object.keys(r)[0];
      const val = String(r[mCol] || '').trim();
      return r._tabType === 'cash' || val.startsWith('7-') || val.startsWith('٧-');
    }).length;
  }, [sheet1, colM1]);

  const payMachinesCount = useMemo(() => {
    return Math.max(0, sheet1.length - cashMachinesCount);
  }, [sheet1, cashMachinesCount]);

  const filteredPreviewRows = useMemo(() => {
    if (sheet1PreviewFilter === 'cash') {
      return sheet1.filter((r) => {
        const mCol = colM1 || Object.keys(r)[0];
        const val = String(r[mCol] || '').trim();
        return r._tabType === 'cash' || val.startsWith('7-') || val.startsWith('٧-');
      });
    }
    if (sheet1PreviewFilter === 'payments') {
      return sheet1.filter((r) => {
        const mCol = colM1 || Object.keys(r)[0];
        const val = String(r[mCol] || '').trim();
        return r._tabType === 'payments' || (!r._tabType && !val.startsWith('7-') && !val.startsWith('٧-'));
      });
    }
    return sheet1;
  }, [sheet1, sheet1PreviewFilter, colM1]);

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
          if (sheet2.length > 0 || sheet3.length > 0) {
            executeReconciliation(data, sheet2, sheet3, colM1 || 'رقم الماكينة', colM2 || 'رقم الماكينة', colAcc, colRep, colM3 || 'رقم الماكينة', colCashAcc, colCashRep, emptyRepFallback);
          }
        }}
        onApplySheet2={(data) => {
          setSheet2(data);
          setSheet2FileName('شيت_المناديب_المعدل_يدوياً');
          showToast(`✅ تم اعتماد ${data.length} سجل مناديب من محرر القوالب.`);
          if (sheet1.length > 0) {
            executeReconciliation(sheet1, data, sheet3, colM1 || 'رقم الماكينة', colM2 || 'رقم الماكينة', colAcc, colRep, colM3 || 'رقم الماكينة', colCashAcc, colCashRep, emptyRepFallback);
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
        sheet3={sheet3}
        colM1={colM1}
        colM2={colM2}
        colM3={colM3}
        colAcc={colAcc}
        colRep={colRep}
        colCashAcc={colCashAcc}
        colCashRep={colCashRep}
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

            {/* 3. قالب المدفوعات (.xlsx) + CSV */}
            <div className="inline-flex items-center bg-[#0d1527] border border-slate-700/80 rounded-xl p-1 text-xs font-bold text-white shrink-0">
              <button
                onClick={() => downloadTemplateFile(2, 'xlsx')}
                className="px-3 py-1.5 flex items-center gap-2 hover:text-emerald-300 transition-colors cursor-pointer"
                title="تنزيل قالب شيت المدفوعات إكسل"
              >
                <FileSpreadsheet className="w-4 h-4 text-emerald-400" />
                <span>قالب المدفوعات (.xlsx)</span>
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

            {/* 4. قالب الكاش (.xlsx) + CSV */}
            <div className="inline-flex items-center bg-[#0d1527] border border-slate-700/80 rounded-xl p-1 text-xs font-bold text-white shrink-0">
              <button
                onClick={() => downloadTemplateFile(3, 'xlsx')}
                className="px-3 py-1.5 flex items-center gap-2 hover:text-emerald-300 transition-colors cursor-pointer"
                title="تنزيل قالب شيت الكاش إكسل"
              >
                <FileSpreadsheet className="w-4 h-4 text-emerald-400" />
                <span>قالب الكاش (.xlsx)</span>
              </button>
              <div className="h-4 w-px bg-slate-700 mx-1"></div>
              <button
                onClick={() => downloadTemplateFile(3, 'csv')}
                className="px-2 py-1 text-[11px] font-black text-blue-400 hover:text-blue-300 transition-colors cursor-pointer"
                title="تنزيل بصيغة CSV"
              >
                CSV
              </button>
            </div>
          </div>
        </div>

        {/* Info & Status Cards */}
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-stretch">
          <div className="lg:col-span-5 bg-gradient-to-br from-slate-900 via-indigo-950 to-slate-950 text-white rounded-3xl p-6 shadow-xl border border-slate-800 flex flex-col justify-between">
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

          {/* User's Exact Smart Control Center: مركز المراقبة الذكي - حالة الجاهزية */}
          <motion.div
            layout
            initial={{ opacity: 0, y: 15 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ type: 'spring', stiffness: 320, damping: 26 }}
            className="lg:col-span-7 flex justify-center w-full"
          >
            <div className="w-full bg-[#0b1120] border border-[#1e293b] rounded-[20px] p-6 sm:p-7 shadow-[0_20px_40px_rgba(0,0,0,0.4)] grid grid-cols-1 md:grid-cols-[1fr_320px] gap-6 items-stretch">
              
              {/* Left Side: File Upload List */}
              <div className="flex flex-col gap-4 justify-center">
                
                {/* File Item 1: شيت الماكينات */}
                <div className={`border rounded-[14px] p-4 sm:px-5 sm:py-4 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 transition-all duration-200 ${
                  sheet1.length > 0 ? 'border-[#059669] bg-[#0f172a]/90 shadow-sm shadow-emerald-500/10' : 'border-[#1e293b] bg-[#0f172a]'
                }`}>
                  <div className="flex items-center gap-3.5">
                    <div className={`w-11 h-11 rounded-[10px] flex items-center justify-center shrink-0 transition-colors ${
                      sheet1.length > 0 ? 'bg-[#10b981]/15 text-[#10b981]' : 'bg-[#1e293b] text-[#94a3b8]'
                    }`}>
                      <FileSpreadsheet className="w-5 h-5" />
                    </div>
                    <div className="flex flex-col gap-1">
                      <div className="flex items-center gap-2">
                        <span className="text-base font-bold text-[#f8fafc]">شيت الماكينات</span>
                        {sheet1.length > 0 && (
                          <span className="text-[10px] font-black text-emerald-400 bg-emerald-500/20 px-2 py-0.5 rounded-full border border-emerald-500/30">
                            ✓ مرفوع
                          </span>
                        )}
                      </div>
                      <span className={`text-xs ${sheet1.length > 0 ? 'text-[#10b981] font-semibold' : 'text-[#64748b]'}`}>
                        {sheet1.length > 0 ? `تم تحميل ${sheet1.length.toLocaleString('ar-EG')} ماكينة بنجاح` : 'لم يتم تحميل الملف بعد'}
                      </span>
                    </div>
                  </div>

                  <div className="flex items-center gap-2 w-full sm:w-auto justify-end">
                    {sheet1.length > 0 && (
                      <button
                        onClick={() => {
                          setSheet1([]);
                          setSheet1FileName('');
                          showToast('🗑️ تم إفراغ شيت الماكينات');
                        }}
                        className="px-2.5 py-1.5 rounded-lg text-xs font-semibold text-rose-400 hover:text-rose-300 bg-rose-500/10 hover:bg-rose-500/20 border border-rose-500/20 transition-all cursor-pointer"
                        title="إفراغ الملف"
                      >
                        إفراغ
                      </button>
                    )}
                    <label className={`px-4 py-2 rounded-lg text-xs font-bold transition-all cursor-pointer shadow-sm text-center ${
                      sheet1.length > 0
                        ? 'bg-[#1e293b] hover:bg-[#334155] text-[#94a3b8] hover:text-white border border-[#334155]'
                        : 'bg-gradient-to-r from-[#4f46e5] to-[#3b82f6] hover:opacity-90 text-white'
                    }`}>
                      <span>{sheet1.length > 0 ? 'تغيير الملف' : 'تحميل الملف'}</span>
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
                    </label>
                  </div>
                </div>

                {/* File Item 2: شيت المدفوعات */}
                <div className={`border rounded-[14px] p-4 sm:px-5 sm:py-4 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 transition-all duration-200 ${
                  sheet2.length > 0 ? 'border-[#059669] bg-[#0f172a]/90 shadow-sm shadow-emerald-500/10' : 'border-[#1e293b] bg-[#0f172a]'
                }`}>
                  <div className="flex items-center gap-3.5">
                    <div className={`w-11 h-11 rounded-[10px] flex items-center justify-center shrink-0 transition-colors ${
                      sheet2.length > 0 ? 'bg-[#10b981]/15 text-[#10b981]' : 'bg-[#1e293b] text-[#94a3b8]'
                    }`}>
                      <Users className="w-5 h-5" />
                    </div>
                    <div className="flex flex-col gap-1">
                      <div className="flex items-center gap-2">
                        <span className="text-base font-bold text-[#f8fafc]">شيت المدفوعات</span>
                        {sheet2.length > 0 && (
                          <span className="text-[10px] font-black text-emerald-400 bg-emerald-500/20 px-2 py-0.5 rounded-full border border-emerald-500/30">
                            ✓ مرفوع
                          </span>
                        )}
                      </div>
                      <span className={`text-xs ${sheet2.length > 0 ? 'text-[#10b981] font-semibold' : 'text-[#64748b]'}`}>
                        {sheet2.length > 0 ? `تم تحميل ${sheet2.length.toLocaleString('ar-EG')} حساب بنجاح` : 'لم يتم تحميل الملف بعد'}
                      </span>
                    </div>
                  </div>

                  <div className="flex items-center gap-2 w-full sm:w-auto justify-end">
                    {sheet2.length > 0 && (
                      <button
                        onClick={() => {
                          setSheet2([]);
                          setSheet2FileName('');
                          showToast('🗑️ تم إفراغ شيت المدفوعات');
                        }}
                        className="px-2.5 py-1.5 rounded-lg text-xs font-semibold text-rose-400 hover:text-rose-300 bg-rose-500/10 hover:bg-rose-500/20 border border-rose-500/20 transition-all cursor-pointer"
                        title="إفراغ الملف"
                      >
                        إفراغ
                      </button>
                    )}
                    <label className={`px-4 py-2 rounded-lg text-xs font-bold transition-all cursor-pointer shadow-sm text-center ${
                      sheet2.length > 0
                        ? 'bg-[#1e293b] hover:bg-[#334155] text-[#94a3b8] hover:text-white border border-[#334155]'
                        : 'bg-gradient-to-r from-[#4f46e5] to-[#3b82f6] hover:opacity-90 text-white'
                    }`}>
                      <span>{sheet2.length > 0 ? 'تغيير الملف' : 'تحميل الملف'}</span>
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
                    </label>
                  </div>
                </div>

                {/* File Item 3: شيت أجهزة الكاش */}
                <div className={`border rounded-[14px] p-4 sm:px-5 sm:py-4 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 transition-all duration-200 ${
                  sheet3.length > 0 ? 'border-[#059669] bg-[#0f172a]/90 shadow-sm shadow-emerald-500/10' : 'border-[#1e293b] bg-[#0f172a]'
                }`}>
                  <div className="flex items-center gap-3.5">
                    <div className={`w-11 h-11 rounded-[10px] flex items-center justify-center shrink-0 transition-colors ${
                      sheet3.length > 0 ? 'bg-[#10b981]/15 text-[#10b981]' : 'bg-[#1e293b] text-[#94a3b8]'
                    }`}>
                      <Calculator className="w-5 h-5" />
                    </div>
                    <div className="flex flex-col gap-1">
                      <div className="flex items-center gap-2">
                        <span className="text-base font-bold text-[#f8fafc]">شيت أجهزة الكاش</span>
                        {sheet3.length > 0 && (
                          <span className="text-[10px] font-black text-emerald-400 bg-emerald-500/20 px-2 py-0.5 rounded-full border border-emerald-500/30">
                            ✓ مرفوع
                          </span>
                        )}
                      </div>
                      <span className={`text-xs ${sheet3.length > 0 ? 'text-[#10b981] font-semibold' : 'text-[#64748b]'}`}>
                        {sheet3.length > 0 ? `تم تحميل ${sheet3.length.toLocaleString('ar-EG')} جهاز بنجاح` : 'لم يتم تحميل الملف بعد'}
                      </span>
                    </div>
                  </div>

                  <div className="flex items-center gap-2 w-full sm:w-auto justify-end">
                    {sheet3.length > 0 && (
                      <button
                        onClick={() => {
                          setSheet3([]);
                          setSheet3FileName('');
                          showToast('🗑️ تم إفراغ شيت الكاش');
                        }}
                        className="px-2.5 py-1.5 rounded-lg text-xs font-semibold text-rose-400 hover:text-rose-300 bg-rose-500/10 hover:bg-rose-500/20 border border-rose-500/20 transition-all cursor-pointer"
                        title="إفراغ الملف"
                      >
                        إفراغ
                      </button>
                    )}
                    <label className={`px-4 py-2 rounded-lg text-xs font-bold transition-all cursor-pointer shadow-sm text-center ${
                      sheet3.length > 0
                        ? 'bg-[#1e293b] hover:bg-[#334155] text-[#94a3b8] hover:text-white border border-[#334155]'
                        : 'bg-gradient-to-r from-[#4f46e5] to-[#3b82f6] hover:opacity-90 text-white'
                    }`}>
                      <span>{sheet3.length > 0 ? 'تغيير الملف' : 'تحميل الملف'}</span>
                      <input
                        type="file"
                        accept=".xlsx, .xls, .csv"
                        className="hidden"
                        onChange={(e) => {
                          if (e.target.files?.length) {
                            handleSheet3Upload(e.target.files[0]);
                          }
                        }}
                      />
                    </label>
                  </div>
                </div>

              </div>

              {/* Right Side: Status Panel */}
              <div className="bg-[#0f172a]/70 border border-[#1e293b] rounded-[16px] p-6 flex flex-col items-center text-center justify-between relative">
                
                <div className="flex flex-col items-center w-full">
                  <div className="bg-indigo-500/15 text-[#818cf8] border border-indigo-500/30 px-4 py-1.5 rounded-full text-xs font-bold mb-3.5 inline-flex items-center gap-1.5">
                    ✨ حالة الجاهزية
                  </div>

                  <h3 className="text-xl sm:text-2xl font-black text-white mb-2">
                    {(sheet1.length > 0 || sheet2.length > 0 || sheet3.length > 0) ? 'جاهز للمطابقة' : 'غير جاهز للمطابقة'}
                  </h3>

                  <p className="text-xs text-[#94a3b8] leading-relaxed mb-6">
                    {(sheet1.length > 0 || sheet2.length > 0 || sheet3.length > 0)
                      ? 'تم التحقق من كافة الملفات المحملة، يمكنك الآن تشغيل نظام المطابقة الذكي.'
                      : 'برجاء تحميل الملفات المطلوبة لبدء تشغيل نظام المطابقة الذكي.'}
                  </p>

                  {/* Progress Box */}
                  <div className="bg-[#0b1120] border border-[#1e293b] rounded-[10px] p-3.5 sm:p-4 w-full mb-6">
                    <div className="flex justify-between text-xs font-bold mb-2.5">
                      <span className="text-[#94a3b8]">الملفات المحملة</span>
                      <span className="text-[#10b981] font-mono">
                        {((sheet1.length > 0 ? 1 : 0) + (sheet2.length > 0 ? 1 : 0) + (sheet3.length > 0 ? 1 : 0))} من 3
                      </span>
                    </div>

                    <div className="bg-[#1e293b] h-2 rounded-full overflow-hidden w-full">
                      <div
                        className="bg-[#10b981] h-full transition-all duration-500 rounded-full"
                        style={{
                          width: `${(((sheet1.length > 0 ? 1 : 0) + (sheet2.length > 0 ? 1 : 0) + (sheet3.length > 0 ? 1 : 0)) / 3) * 100}%`,
                        }}
                      />
                    </div>
                  </div>
                </div>

                {/* Main Run Button */}
                <button
                  onClick={runReconciliation}
                  className="w-full py-3.5 border-none rounded-[10px] bg-gradient-to-r from-[#10b981] to-[#3b82f6] hover:opacity-95 active:scale-[0.98] text-white text-base font-extrabold cursor-pointer flex items-center justify-center gap-2 shadow-lg shadow-emerald-500/20 transition-all"
                >
                  <Zap className="w-5 h-5 text-amber-300" />
                  <span>تشغيل المطابقة الآن</span>
                </button>

              </div>

            </div>
          </motion.div>
        </div>

        {/* Navigation Tabs - Modern Cards Grid */}
        <div className="grid grid-cols-2 sm:grid-cols-4 xl:grid-cols-7 gap-2.5 sm:gap-3">
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
                2. شيت المدفوعات
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

          {/* Card 3: Sheet 3 */}
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
                <Calculator className="w-4 h-4" />
              </div>
            </div>

            <div>
              <h4
                className={`text-xs font-bold leading-tight ${
                  currentTab === 3 ? 'text-white' : 'text-slate-900 dark:text-white'
                }`}
              >
                3. شيت الكاش
              </h4>
              <p
                className={`text-[10px] mt-1 font-medium truncate ${
                  currentTab === 3 ? 'text-indigo-100 font-bold' : 'text-slate-500 dark:text-slate-400'
                }`}
              >
                {sheet3.length > 0 ? `${sheet3.length.toLocaleString('ar-EG')} سجل كاش` : 'ربط حسابات الكاش'}
              </p>
            </div>

            {currentTab === 3 && (
              <div className="absolute bottom-0 left-0 right-0 h-1 bg-white/70 rounded-full mx-3 mb-0.5" />
            )}
          </button>

          {/* Card 4: Mapping */}
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
                    : 'bg-amber-500/10 text-amber-500 dark:text-amber-400 group-hover:bg-amber-500/20'
                }`}
              >
                <Sliders className="w-4 h-4" />
              </div>
            </div>

            <div>
              <h4
                className={`text-xs font-bold leading-tight ${
                  currentTab === 4 ? 'text-white' : 'text-slate-900 dark:text-white'
                }`}
              >
                4. تعيين الأعمدة
              </h4>
              <p
                className={`text-[10px] mt-1 font-medium truncate ${
                  currentTab === 4 ? 'text-indigo-100 font-bold' : 'text-slate-500 dark:text-slate-400'
                }`}
              >
                {colM1 && colM2 ? 'أعمدة الربط محددة' : 'ضبط ومطابقة الأعمدة'}
              </p>
            </div>

            {currentTab === 4 && (
              <div className="absolute bottom-0 left-0 right-0 h-1 bg-white/70 rounded-full mx-3 mb-0.5" />
            )}
          </button>

          {/* Card 5: Reports */}
          <button
            onClick={() => setCurrentTab(5)}
            className={`p-3 sm:p-3.5 rounded-2xl border text-right transition-all duration-200 flex flex-col justify-between gap-2.5 cursor-pointer group relative overflow-hidden ${
              currentTab === 5
                ? 'bg-indigo-600 border-indigo-400 text-white shadow-lg shadow-indigo-600/30 ring-2 ring-indigo-400/40'
                : 'bg-white dark:bg-[#0a0f1d] border-slate-200 dark:border-slate-800 text-slate-700 dark:text-slate-300 hover:border-indigo-400/60 hover:bg-slate-50 dark:hover:bg-[#0e1628]'
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
                    : 'bg-indigo-500/10 text-indigo-500 dark:text-indigo-400 group-hover:bg-indigo-500/20'
                }`}
              >
                <Layers className="w-4 h-4" />
              </div>
            </div>

            <div>
              <h4
                className={`text-xs font-bold leading-tight ${
                  currentTab === 5 ? 'text-white' : 'text-slate-900 dark:text-white'
                }`}
              >
                5. تقرير الأسطر
              </h4>
              <p
                className={`text-[10px] mt-1 font-mono font-bold truncate ${
                  currentTab === 5 ? 'text-indigo-100' : 'text-indigo-600 dark:text-indigo-400'
                }`}
              >
                {expandedRows.length.toLocaleString('ar-EG')} سطر مفصل
              </p>
            </div>

            {currentTab === 5 && (
              <div className="absolute bottom-0 left-0 right-0 h-1 bg-white/70 rounded-full mx-3 mb-0.5" />
            )}
          </button>

          {/* Card 6: Audit & Conflicts */}
          <button
            onClick={() => setCurrentTab(6)}
            className={`p-3 sm:p-3.5 rounded-2xl border text-right transition-all duration-200 flex flex-col justify-between gap-2.5 cursor-pointer group relative overflow-hidden ${
              currentTab === 6
                ? 'bg-amber-600 border-amber-400 text-white shadow-lg shadow-amber-600/30 ring-2 ring-amber-400/40'
                : 'bg-white dark:bg-[#0a0f1d] border-slate-200 dark:border-slate-800 text-slate-700 dark:text-slate-300 hover:border-amber-400/60 hover:bg-slate-50 dark:hover:bg-[#0e1628]'
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
                    : 'bg-amber-500/10 text-amber-500 dark:text-amber-400 group-hover:bg-amber-500/20'
                }`}
              >
                <CircleAlert className="w-4 h-4" />
              </div>
            </div>

            <div>
              <h4
                className={`text-xs font-bold leading-tight ${
                  currentTab === 6 ? 'text-white' : 'text-amber-600 dark:text-amber-400'
                }`}
              >
                6. فحص النزاعات
              </h4>
              <p
                className={`text-[10px] mt-1 font-mono font-bold truncate ${
                  currentTab === 6 ? 'text-amber-100' : 'text-slate-500 dark:text-slate-400'
                }`}
              >
                {duplicates.length + irregulars.length > 0
                  ? `${(duplicates.length + irregulars.length).toLocaleString('ar-EG')} شاذ / مكرر`
                  : 'فحص سليم (0)'}
              </p>
            </div>

            {currentTab === 6 && (
              <div className="absolute bottom-0 left-0 right-0 h-1 bg-white/70 rounded-full mx-3 mb-0.5" />
            )}
          </button>

          {/* Card 7: Desktop App Studio */}
          <button
            onClick={() => setCurrentTab(7)}
            className={`p-3 sm:p-3.5 rounded-2xl border text-right transition-all duration-200 flex flex-col justify-between gap-2.5 cursor-pointer group relative overflow-hidden ${
              currentTab === 7
                ? 'bg-blue-600 border-blue-400 text-white shadow-lg shadow-blue-600/30 ring-2 ring-blue-400/40'
                : 'bg-white dark:bg-[#0a0f1d] border-slate-200 dark:border-slate-800 text-slate-700 dark:text-slate-300 hover:border-blue-400/60 hover:bg-slate-50 dark:hover:bg-[#0e1628]'
            }`}
          >
            <div className="flex items-center justify-between w-full">
              <span
                className={`text-[10px] font-mono font-black px-2 py-0.5 rounded-md ${
                  currentTab === 7
                    ? 'bg-white/20 text-white'
                    : 'bg-slate-100 dark:bg-slate-800 text-slate-500 dark:text-slate-400'
                }`}
              >
                07
              </span>
              <div
                className={`w-7 h-7 rounded-xl flex items-center justify-center transition-colors ${
                  currentTab === 7
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
                    currentTab === 7 ? 'text-white' : 'text-slate-900 dark:text-white'
                  }`}
                >
                  7. استوديو HTML
                </h4>
                <span className="w-2 h-2 rounded-full bg-emerald-400 animate-pulse shrink-0" title="Live Update" />
              </div>
              <p
                className={`text-[10px] mt-1 font-medium truncate ${
                  currentTab === 7 ? 'text-blue-100 font-bold' : 'text-slate-500 dark:text-slate-400'
                }`}
              >
                تحديث أوفلاين
              </p>
            </div>

            {currentTab === 7 && (
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

                <div className="flex items-center gap-2 flex-wrap">
                  <button
                    onClick={handleApplyCashPrefix}
                    className="px-3.5 py-1.5 text-xs font-black rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white flex items-center gap-1.5 transition-all shadow-md shadow-emerald-600/20 cursor-pointer"
                    title="تحويل كافة ماكينات الكاش في الشيت المرفوع تلقائياً إلى البادئة 7-"
                  >
                    <Zap className="w-3.5 h-3.5 text-amber-300" />
                    <span>⚡ تحويل ماكينات الكاش لـ 7-</span>
                  </button>
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
                <div className="mt-5 space-y-4 border-t border-slate-200 dark:border-slate-800/80 pt-5">
                  <div className="p-3.5 rounded-2xl bg-emerald-500/10 border border-emerald-500/20 text-emerald-600 dark:text-emerald-400 text-xs font-bold flex flex-wrap items-center justify-between gap-3">
                    <div className="flex items-center gap-2">
                      <CheckCircle2 className="w-4 h-4 text-emerald-500 shrink-0" />
                      <span>تم قراءة شيت الماكينات بنجاح: {sheet1.length.toLocaleString('ar-EG')} ماكينة</span>
                      {sheet1FileName && <span className="text-slate-400 font-normal">({sheet1FileName})</span>}
                    </div>

                    <div className="flex items-center gap-2">
                      <button
                        onClick={handleApplyCashPrefix}
                        className="px-3 py-1 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg text-xs font-black transition cursor-pointer flex items-center gap-1 shadow-sm"
                      >
                        <Zap className="w-3.5 h-3.5 text-amber-300" />
                        <span>⚡ تحويل ماكينات الكاش لـ 7-</span>
                      </button>
                      <button
                        onClick={() => setCurrentTab(2)}
                        className="text-indigo-600 dark:text-indigo-400 hover:text-indigo-500 underline font-bold cursor-pointer transition text-xs"
                      >
                        الانتقال لرفع شيت المناديب &larr;
                      </button>
                    </div>
                  </div>

                  {/* Live Sheet 1 Preview Table with Interactive Filter Pills */}
                  <div className="bg-slate-50 dark:bg-slate-800/40 rounded-2xl border border-slate-200 dark:border-slate-800 overflow-hidden shadow-sm">
                    <div className="p-3.5 bg-slate-100 dark:bg-slate-800 flex flex-wrap items-center justify-between gap-3 border-b border-slate-200 dark:border-slate-700">
                      <div className="flex items-center gap-2">
                        <Eye className="w-4 h-4 text-indigo-500" />
                        <h4 className="text-xs font-bold text-slate-800 dark:text-slate-200">
                          معاينة بيانات شيت الماكينات المرفوع مباشرة ({Math.min(25, filteredPreviewRows.length)} من {filteredPreviewRows.length.toLocaleString('ar-EG')})
                        </h4>
                      </div>

                      {/* Filter Pills for Cash & Payments */}
                      <div className="flex items-center gap-1.5 bg-white dark:bg-slate-900 p-1 rounded-xl border border-slate-200 dark:border-slate-700">
                        <button
                          onClick={() => setSheet1PreviewFilter('all')}
                          className={`px-2.5 py-1 rounded-lg text-[11px] font-bold transition-all cursor-pointer ${
                            sheet1PreviewFilter === 'all'
                              ? 'bg-indigo-600 text-white shadow-sm'
                              : 'text-slate-600 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-800'
                          }`}
                        >
                          كافة الماكينات ({sheet1.length})
                        </button>

                        <button
                          onClick={() => setSheet1PreviewFilter('cash')}
                          className={`px-2.5 py-1 rounded-lg text-[11px] font-bold flex items-center gap-1 transition-all cursor-pointer ${
                            sheet1PreviewFilter === 'cash'
                              ? 'bg-emerald-600 text-white shadow-sm'
                              : 'text-slate-600 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-800'
                          }`}
                        >
                          <span>💵 ماكينات الكاش</span>
                          <span className="px-1.5 py-0.2 rounded-full bg-emerald-950/40 text-emerald-200 font-mono text-[10px]">
                            {cashMachinesCount}
                          </span>
                        </button>

                        <button
                          onClick={() => setSheet1PreviewFilter('payments')}
                          className={`px-2.5 py-1 rounded-lg text-[11px] font-bold flex items-center gap-1 transition-all cursor-pointer ${
                            sheet1PreviewFilter === 'payments'
                              ? 'bg-blue-600 text-white shadow-sm'
                              : 'text-slate-600 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-800'
                          }`}
                        >
                          <span>💳 ماكينات المدفوعات</span>
                          <span className="px-1.5 py-0.2 rounded-full bg-blue-950/40 text-blue-200 font-mono text-[10px]">
                            {payMachinesCount}
                          </span>
                        </button>
                      </div>
                    </div>

                    <div className="overflow-x-auto max-h-80 overflow-y-auto">
                      <table className="w-full text-right text-xs">
                        <thead className="bg-slate-200/80 dark:bg-slate-900 text-slate-700 dark:text-slate-300 font-bold sticky top-0 border-b border-slate-200 dark:border-slate-800 z-10">
                          <tr>
                            <th className="p-2.5 w-12 text-center font-mono">م</th>
                            {Object.keys(sheet1[0]).map((col) => (
                              <th key={col} className="p-2.5">{col}</th>
                            ))}
                            <th className="p-2.5 text-center">تصنيف الماكينة</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-200 dark:divide-slate-800/60 font-mono">
                          {filteredPreviewRows.length === 0 ? (
                            <tr>
                              <td colSpan={Object.keys(sheet1[0]).length + 2} className="p-8 text-center text-slate-400">
                                لا توجد ماكينات مطابقة لهذا التصنيف في الشيت المرفوع حالياً.
                              </td>
                            </tr>
                          ) : (
                            filteredPreviewRows.slice(0, 25).map((row, idx) => {
                              const mCol = colM1 || Object.keys(row)[0];
                              const machineVal = String(row[mCol] || '').trim();
                              const isCash = machineVal.startsWith('7-') || machineVal.startsWith('٧-');
                              return (
                                <tr key={idx} className="hover:bg-slate-100 dark:hover:bg-slate-800/60 transition-colors">
                                  <td className="p-2 text-center text-slate-400 font-bold">{idx + 1}</td>
                                  {Object.keys(sheet1[0]).map((col) => (
                                    <td key={col} className="p-2 font-bold text-slate-800 dark:text-slate-200">
                                      {String(row[col] ?? '')}
                                    </td>
                                  ))}
                                  <td className="p-2 text-center">
                                    {isCash ? (
                                      <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-black bg-emerald-500/10 text-emerald-600 dark:text-emerald-400 border border-emerald-500/20">
                                        💵 ماكينة كاش (7-)
                                      </span>
                                    ) : (
                                      <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-black bg-blue-500/10 text-blue-600 dark:text-blue-400 border border-blue-500/20">
                                        💳 مدفوعات
                                      </span>
                                    )}
                                  </td>
                                </tr>
                              );
                            })
                          )}
                        </tbody>
                      </table>
                    </div>

                    {filteredPreviewRows.length > 25 && (
                      <div className="p-2.5 bg-slate-100/50 dark:bg-slate-800/50 text-center text-[11px] text-slate-500 dark:text-slate-400 border-t border-slate-200 dark:border-slate-700 font-semibold">
                        عرض أول 25 ماكينة من إجمالي {filteredPreviewRows.length.toLocaleString('ar-EG')} ماكينة محملة في هذا التصنيف.
                      </div>
                    )}
                  </div>
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
                    <span>رفع الشيت الثاني (شيت مدفوعات المناديب وأرقام حساباتهم)</span>
                  </h3>
                  <p className="text-xs text-slate-500 dark:text-slate-400">
                    الملف الذي يحتوي على رقم الماكينة، اسم المندوب، ورقم حساب المندوب للمدفوعات (Account No / Code).
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
                  اضغط هنا لاختيار شيت المدفوعات أو اسحبه وأفلته هنا
                </p>
                <p className="text-[11px] text-slate-400">يدعم ملفات Excel (.xlsx, .xls) أو CSV</p>
              </label>

              {sheet2.length > 0 && (
                <div className="mt-5 space-y-4 border-t border-slate-200 dark:border-slate-800/80 pt-5">
                  <div className="p-3.5 rounded-2xl bg-emerald-500/10 border border-emerald-500/20 text-emerald-600 dark:text-emerald-400 text-xs font-bold flex items-center justify-between">
                    <div className="flex items-center gap-2">
                      <CheckCircle2 className="w-4 h-4 text-emerald-500 shrink-0" />
                      <span>تم قراءة شيت المدفوعات بنجاح: {sheet2.length.toLocaleString('ar-EG')} سجل توزيع</span>
                      {sheet2FileName && <span className="text-slate-400 font-normal">({sheet2FileName})</span>}
                    </div>
                    <button
                      onClick={() => setCurrentTab(3)}
                      className="text-indigo-600 dark:text-indigo-400 underline font-bold cursor-pointer transition text-xs"
                    >
                      الانتقال لرفع شيت الكاش &larr;
                    </button>
                  </div>

                  {/* Live Sheet 2 Preview Table */}
                  <div className="bg-slate-50 dark:bg-slate-800/40 rounded-2xl border border-slate-200 dark:border-slate-800 overflow-hidden shadow-sm">
                    <div className="p-3.5 bg-slate-100 dark:bg-slate-800 flex items-center justify-between border-b border-slate-200 dark:border-slate-700">
                      <div className="flex items-center gap-2">
                        <Eye className="w-4 h-4 text-blue-500" />
                        <h4 className="text-xs font-bold text-slate-800 dark:text-slate-200">
                          معاينة بيانات شيت المدفوعات المرفوع مباشرة ({Math.min(25, sheet2.length)} من {sheet2.length.toLocaleString('ar-EG')})
                        </h4>
                      </div>
                      <span className="text-[10px] font-bold text-slate-500 dark:text-slate-400 bg-white dark:bg-slate-900 px-2 py-0.5 rounded-md border border-slate-200 dark:border-slate-700">
                        معاينة الشيت أوفلاين 100%
                      </span>
                    </div>

                    <div className="overflow-x-auto max-h-80 overflow-y-auto">
                      <table className="w-full text-right text-xs">
                        <thead className="bg-slate-200/80 dark:bg-slate-900 text-slate-700 dark:text-slate-300 font-bold sticky top-0 border-b border-slate-200 dark:border-slate-800 z-10">
                          <tr>
                            <th className="p-2.5 w-12 text-center font-mono">م</th>
                            {Object.keys(sheet2[0]).map((col) => (
                              <th key={col} className="p-2.5">{col}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-200 dark:divide-slate-800/60 font-mono">
                          {sheet2.slice(0, 25).map((row, idx) => (
                            <tr key={idx} className="hover:bg-slate-100 dark:hover:bg-slate-800/60 transition-colors">
                              <td className="p-2 text-center text-slate-400 font-bold">{idx + 1}</td>
                              {Object.keys(sheet2[0]).map((col) => (
                                <td key={col} className="p-2 font-bold text-slate-800 dark:text-slate-200">
                                  {String(row[col] ?? '')}
                                </td>
                              ))}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>

                    {sheet2.length > 25 && (
                      <div className="p-2.5 bg-slate-100/50 dark:bg-slate-800/50 text-center text-[11px] text-slate-500 dark:text-slate-400 border-t border-slate-200 dark:border-slate-700 font-semibold">
                        عرض أول 25 سجل من إجمالي {sheet2.length.toLocaleString('ar-EG')} سجل محمل في الشيت.
                      </div>
                    )}
                  </div>
                </div>
              )}
            </div>
          </div>
        )}

        {/* TAB 3: Sheet 3 Upload (Cash) */}
        {currentTab === 3 && (
          <div className="space-y-4">
            <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-3xl p-6 sm:p-8">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-4">
                <div>
                  <h3 className="text-sm font-bold text-slate-900 dark:text-white mb-1 flex items-center gap-2">
                    <Calculator className="w-4 h-4 text-emerald-500" />
                    <span>رفع الشيت الثالث (شيت الكاش وحسابات المناديب)</span>
                  </h3>
                  <p className="text-xs text-slate-500 dark:text-slate-400">
                    الملف الذي يحتوي على رقم الماكينة، اسم المندوب المسؤول عن الكاش، ورقم حساب الكاش (Account No / Code).
                  </p>
                </div>

                <div className="flex items-center gap-2">
                  <button
                    onClick={() => downloadTemplateFile(3, 'xlsx')}
                    className="px-3 py-1.5 text-xs font-bold rounded-xl bg-emerald-50 hover:bg-emerald-100 dark:bg-emerald-950/60 text-emerald-600 dark:text-emerald-300 border border-emerald-200 dark:border-emerald-800 flex items-center gap-1.5 transition-colors cursor-pointer"
                  >
                    <Download className="w-3.5 h-3.5" />
                    <span>تحميل نموذج القالب</span>
                  </button>
                </div>
              </div>

              {/* Dropzone */}
              <label className="relative border-2 border-dashed border-slate-300 dark:border-slate-700 rounded-3xl p-8 text-center cursor-pointer hover:border-emerald-500 bg-slate-50 dark:bg-slate-800/30 block transition-colors">
                <input
                  type="file"
                  accept=".xlsx, .xls, .csv"
                  className="hidden"
                  onChange={(e) => {
                    if (e.target.files?.length) {
                      handleSheet3Upload(e.target.files[0]);
                    }
                  }}
                />
                <FileSpreadsheet className="w-12 h-12 text-emerald-500 mx-auto mb-2" />
                <p className="text-xs font-bold text-slate-800 dark:text-slate-200 mb-1">
                  اضغط هنا لاختيار شيت الكاش أو اسحبه وأفلته هنا
                </p>
                <p className="text-[11px] text-slate-400">يدعم ملفات Excel (.xlsx, .xls) أو CSV</p>
              </label>

              {sheet3.length > 0 && (
                <div className="mt-6 p-4 rounded-2xl bg-[#080d16]/60 border border-slate-850 space-y-4">
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 bg-emerald-500/10 border border-emerald-500/20 p-3.5 rounded-xl text-emerald-600 dark:text-emerald-400 text-xs font-bold">
                    <span>تم قراءة شيت الكاش بنجاح: {sheet3.length.toLocaleString('ar-EG')} سجل كاش</span>
                    <div className="flex items-center gap-3">
                      <button
                        onClick={handleApplyCashPrefix}
                        className="px-3 py-1 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg text-[10px] font-black transition cursor-pointer"
                      >
                        ⚡ تحويل لـ 7-
                      </button>
                      <button
                        onClick={() => setCurrentTab(4)}
                        className="text-indigo-600 dark:text-indigo-400 hover:text-indigo-500 underline font-bold cursor-pointer transition"
                      >
                        الانتقال لتعيين الأعمدة &larr;
                      </button>
                    </div>
                  </div>

                  {/* Smart Helper Card */}
                  <div className="p-4 rounded-xl bg-slate-800/20 border border-slate-800 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                    <div className="space-y-1">
                      <h4 className="text-xs font-bold text-slate-200 flex items-center gap-1.5">
                        <span>💡 منسق أرقام ماكينات الكاش السريع</span>
                      </h4>
                      <p className="text-[11px] text-slate-400">
                        إذا كانت أرقام الماكينات في شيت الكاش المرفوع تبدأ بأرقام عادية (مثل <strong className="text-amber-400">123</strong>) وتريد تحويلها فوراً للتنسيق المطلوب المعتمد المقترن ببادئة الكاش (لتصبح <strong className="text-emerald-400">7-123</strong>) بضغطة زر واحدة:
                      </p>
                    </div>

                    <motion.button
                      whileHover={{ scale: 1.03 }}
                      whileTap={{ scale: 0.97 }}
                      transition={{ type: 'spring', stiffness: 450, damping: 22 }}
                      onClick={handleApplyCashPrefix}
                      className="px-4 py-2.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-xs font-black transition-all shadow-lg shadow-emerald-600/30 flex items-center gap-1.5 shrink-0 cursor-pointer"
                    >
                      <Zap className="w-4 h-4 text-amber-300" />
                      <span>⚡ تحويل ماكينات الكاش إلى 7-</span>
                    </motion.button>
                  </div>

                  {/* Live Sheet 3 Preview Table */}
                  <div className="bg-slate-50 dark:bg-slate-800/40 rounded-2xl border border-slate-200 dark:border-slate-800 overflow-hidden shadow-sm">
                    <div className="p-3.5 bg-slate-100 dark:bg-slate-800 flex items-center justify-between border-b border-slate-200 dark:border-slate-700">
                      <div className="flex items-center gap-2">
                        <Eye className="w-4 h-4 text-emerald-500" />
                        <h4 className="text-xs font-bold text-slate-800 dark:text-slate-200">
                          معاينة بيانات شيت الكاش المرفوع مباشرة ({Math.min(25, sheet3.length)} من {sheet3.length.toLocaleString('ar-EG')})
                        </h4>
                      </div>
                      <span className="text-[10px] font-bold text-slate-500 dark:text-slate-400 bg-white dark:bg-slate-900 px-2 py-0.5 rounded-md border border-slate-200 dark:border-slate-700">
                        معاينة الشيت أوفلاين 100%
                      </span>
                    </div>

                    <div className="overflow-x-auto max-h-80 overflow-y-auto">
                      <table className="w-full text-right text-xs">
                        <thead className="bg-slate-200/80 dark:bg-slate-900 text-slate-700 dark:text-slate-300 font-bold sticky top-0 border-b border-slate-200 dark:border-slate-800 z-10">
                          <tr>
                            <th className="p-2.5 w-12 text-center font-mono">م</th>
                            {Object.keys(sheet3[0]).map((col) => (
                              <th key={col} className="p-2.5">{col}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-200 dark:divide-slate-800/60 font-mono">
                          {sheet3.slice(0, 25).map((row, idx) => (
                            <tr key={idx} className="hover:bg-slate-100 dark:hover:bg-slate-800/60 transition-colors">
                              <td className="p-2 text-center text-slate-400 font-bold">{idx + 1}</td>
                              {Object.keys(sheet3[0]).map((col) => (
                                <td key={col} className="p-2 font-bold text-slate-800 dark:text-slate-200">
                                  {String(row[col] ?? '')}
                                </td>
                              ))}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>

                    {sheet3.length > 25 && (
                      <div className="p-2.5 bg-slate-100/50 dark:bg-slate-800/50 text-center text-[11px] text-slate-500 dark:text-slate-400 border-t border-slate-200 dark:border-slate-700 font-semibold">
                        عرض أول 25 سجل من إجمالي {sheet3.length.toLocaleString('ar-EG')} سجل كاش محمل في الشيت.
                      </div>
                    )}
                  </div>
                </div>
              )}
            </div>
          </div>
        )}



        {/* TAB 4: Column Mapping */}
        {currentTab === 4 && (
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
              <h4 className="text-xs font-bold text-[#4f46e5] dark:text-[#818cf8] mb-2">تعيين أعمدة المدفوعات والماكينات (شيت 1 وشيت 2)</h4>
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 mb-6">
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
                    2. عمود الماكينة (شيت المدفوعات):
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
                      <option value="">-- يرجى رفع شيت المدفوعات أولاً --</option>
                    )}
                  </select>
                  <p className="text-[10px] text-slate-400">رقم الماكينة المقابل في شيت المدفوعات</p>
                </div>

                {/* Account Col */}
                <div className="p-4 rounded-2xl bg-slate-50 dark:bg-slate-800/40 border border-slate-200 dark:border-slate-800 space-y-2">
                  <label className="block text-xs font-bold text-amber-600 dark:text-amber-400">
                    3. عمود رقم حساب المدفوعات:
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
                  <p className="text-[10px] text-slate-400">كود / حساب مسؤول المدفوعات</p>
                </div>

                {/* Rep Name Col */}
                <div className="p-4 rounded-2xl bg-slate-50 dark:bg-slate-800/40 border border-slate-200 dark:border-slate-800 space-y-2">
                  <label className="block text-xs font-bold text-emerald-600 dark:text-emerald-400">
                    4. عمود اسم مندوب المدفوعات:
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
                      <option value="">-- يرجى رفع شيت المدفوعات أولاً --</option>
                    )}
                  </select>
                  <p className="text-[10px] text-slate-400">عمود اسم المندوب في شيت المدفوعات</p>
                </div>
              </div>

              <h4 className="text-xs font-bold text-emerald-600 dark:text-emerald-400 mb-2">تعيين أعمدة الكاش (شيت 3)</h4>
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
                {/* Machine Col 3 */}
                <div className="p-4 rounded-2xl bg-slate-50 dark:bg-slate-800/40 border border-slate-200 dark:border-slate-800 space-y-2">
                  <label className="block text-xs font-bold text-indigo-600 dark:text-indigo-400">
                    5. عمود الماكينة (شيت الكاش):
                  </label>
                  <select
                    className="w-full p-2 bg-white dark:bg-slate-900 border border-slate-300 dark:border-slate-700 rounded-xl text-xs font-semibold focus:border-indigo-500 focus:outline-none"
                    value={colM3}
                    onChange={(e) => setColM3(e.target.value)}
                  >
                    {sheet3.length > 0 ? (
                      Object.keys(sheet3[0]).map((c) => (
                        <option key={c} value={c}>
                          {c}
                        </option>
                      ))
                    ) : (
                      <option value="">-- يرجى رفع شيت الكاش أولاً --</option>
                    )}
                  </select>
                  <p className="text-[10px] text-slate-400">رقم الماكينة المقابل في شيت الكاش</p>
                </div>

                {/* Cash Account Col */}
                <div className="p-4 rounded-2xl bg-slate-50 dark:bg-slate-800/40 border border-slate-200 dark:border-slate-800 space-y-2">
                  <label className="block text-xs font-bold text-amber-600 dark:text-amber-400">
                    6. عمود رقم حساب الكاش:
                  </label>
                  <select
                    className="w-full p-2 bg-white dark:bg-slate-900 border border-slate-300 dark:border-slate-700 rounded-xl text-xs font-semibold focus:border-amber-500 focus:outline-none"
                    value={colCashAcc}
                    onChange={(e) => setColCashAcc(e.target.value)}
                  >
                    <option value="">-- بدون عمود حساب كاش --</option>
                    {sheet3.length > 0 &&
                      Object.keys(sheet3[0]).map((c) => (
                        <option key={c} value={c}>
                          {c}
                        </option>
                      ))}
                  </select>
                  <p className="text-[10px] text-slate-400">كود / حساب مسؤول الكاش</p>
                </div>

                {/* Cash Rep Name Col */}
                <div className="p-4 rounded-2xl bg-slate-50 dark:bg-slate-800/40 border border-slate-200 dark:border-slate-800 space-y-2">
                  <label className="block text-xs font-bold text-emerald-600 dark:text-emerald-400">
                    7. عمود اسم مندوب الكاش:
                  </label>
                  <select
                    className="w-full p-2 bg-white dark:bg-slate-900 border border-slate-300 dark:border-slate-700 rounded-xl text-xs font-semibold focus:border-emerald-500 focus:outline-none"
                    value={colCashRep}
                    onChange={(e) => setColCashRep(e.target.value)}
                  >
                    {sheet3.length > 0 ? (
                      Object.keys(sheet3[0]).map((c) => (
                        <option key={c} value={c}>
                          {c}
                        </option>
                      ))
                    ) : (
                      <option value="">-- يرجى رفع شيت الكاش أولاً --</option>
                    )}
                  </select>
                  <p className="text-[10px] text-slate-400">عمود اسم المندوب في شيت الكاش</p>
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
                  onClick={() => setCurrentTab(3)}
                  className="text-xs font-bold text-slate-500 hover:text-slate-700 dark:hover:text-slate-300 flex items-center gap-1.5 cursor-pointer"
                >
                  <ArrowRight className="w-4 h-4" />
                  <span>الرجوع لشيت الكاش</span>
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

        {/* TAB 5: Detailed Reconciliation Reports & Sequential Rows */}
        {currentTab === 5 && (
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

                  {/* Excel Export & Print Buttons */}
                  <div className="flex items-center gap-2">
                    <button
                      onClick={() => window.print()}
                      className="px-3.5 py-2 rounded-xl bg-blue-600 hover:bg-blue-500 text-white text-xs font-bold flex items-center gap-1.5 shadow-md transition-colors cursor-pointer shrink-0"
                      title="طباعة التقرير أو حفظه كملف PDF"
                    >
                      <Download className="w-4 h-4 text-blue-200" />
                      <span>طباعة / PDF</span>
                    </button>
                    <button
                      onClick={handleExportExcel}
                      className="px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold flex items-center gap-2 shadow-md transition-colors cursor-pointer shrink-0"
                    >
                      <FileSpreadsheet className="w-4 h-4 text-emerald-200" />
                      <span>تصدير Excel (أسطر منفصلة)</span>
                    </button>
                  </div>
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
                                  <div className="flex flex-col items-end gap-1 shrink-0">
                                    <span
                                      className={`px-2 py-0.5 rounded-xl text-[10px] font-black border ${
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
                                    <span
                                      className={`px-2 py-0.5 rounded-xl text-[10px] font-black border ${
                                        item.type === 'cash'
                                          ? 'bg-emerald-500/20 text-emerald-400 border-emerald-500/30'
                                          : 'bg-blue-500/20 text-blue-400 border-blue-500/30'
                                      }`}
                                    >
                                      {item.type === 'cash' ? '💵 كاش' : '💳 مدفوعات'}
                                    </span>
                                  </div>
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
                        <th className="p-3.5 w-32 text-center">نوع الحساب</th>
                        <th className="p-3.5 w-40">رقم حساب المندوب</th>
                        <th className="p-3.5">اسم المندوب المسند</th>
                        <th className="p-3.5 w-48 text-center">ترتيب المندوب للماكينة</th>
                        <th className="p-3.5 w-40 text-center">حالة الماكينة</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-800/80 text-slate-300">
                      {displayedSlice.length === 0 ? (
                        <tr>
                          <td colSpan={8} className="p-8 text-center text-slate-400">
                            لا توجد نتائج تطابق شرط البحث أو الفلتر. ارفع الشيتات أو اختر «عينة فورية».
                          </td>
                        </tr>
                      ) : (
                        <>
                          {tableVirtualState.topPadding > 0 && (
                            <tr style={{ border: 'none' }}>
                              <td colSpan={8} style={{ height: tableVirtualState.topPadding, padding: 0, border: 'none' }} />
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

                                {/* Account Type Column */}
                                 <td className="p-3.5 text-center font-bold">
                                   {item.type === 'cash' ? (
                                     <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[11px] font-black bg-emerald-500/10 text-emerald-400 border border-emerald-500/20">
                                       💵 كاش
                                     </span>
                                   ) : (
                                     <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[11px] font-black bg-blue-500/10 text-blue-400 border border-blue-500/20">
                                       💳 مدفوعات
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
                              <td colSpan={8} style={{ height: tableVirtualState.bottomPadding, padding: 0, border: 'none' }} />
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

        {/* TAB 6: Audit & Conflict Inspection */}
        {currentTab === 6 && (
          <AuditConflictsTab
            duplicates={duplicates}
            irregulars={irregulars}
            onFilterDuplicates={() => {
              if (duplicates.length > 0) {
                setSearchQuery(duplicates[0].machine);
                setCurrentFilter('all');
                setCurrentTab(5);
                showToast(`🔍 تم تصفية الماكينة المكررة: ${duplicates[0].machine}`);
              }
            }}
            onFilterIrregulars={() => {
              if (irregulars.length > 0) {
                setSearchQuery(irregulars[0].account);
                setCurrentFilter('all');
                setCurrentTab(5);
                showToast(`🔍 تم تصفية الحساب الشاذ: ${irregulars[0].account}`);
              }
            }}
          />
        )}

        {/* TAB 7: Standalone Desktop App & HTML Exporter */}
        {currentTab === 7 && (
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
