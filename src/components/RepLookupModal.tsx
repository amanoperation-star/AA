import React, { useState, useMemo, useEffect } from 'react';
import {
  Search,
  User,
  Layers,
  FileSpreadsheet,
  X,
  CheckCircle2,
  AlertCircle,
  Users,
  HardDrive,
  Filter,
  FileCheck,
  Smartphone,
  Cpu,
  ArrowRightLeft,
  Sparkles,
  HelpCircle,
  TableProperties,
  Repeat,
  AlertTriangle,
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { MachineSummary, ExpandedRow, SheetRow } from '../types';

function normalizeMachineId(id: string | number | undefined | null): string {
  if (id === undefined || id === null) return '';
  let clean = String(id).replace(/[\u200B-\u200D\uFEFF\u00A0]/g, '').trim();
  // Convert Eastern Arabic digits (٠-٩) to Western digits (0-9)
  clean = clean.replace(/[٠-٩]/g, (d) => String('٠١٢٣٤٥٦٧٨٩'.indexOf(d)));
  // Remove "7-" or "٧-" prefixes
  if (clean.startsWith('7-')) {
    clean = clean.substring(2).trim();
  } else if (clean.startsWith('٧-')) {
    clean = clean.substring(2).trim();
  }
  // Remove trailing .0 from float string representation (e.g. 1001.0 -> 1001)
  if (clean.endsWith('.0')) {
    clean = clean.substring(0, clean.length - 2);
  }
  return clean;
}

function normalizeText(text: string | number | undefined | null): string {
  if (text === undefined || text === null) return '';
  let clean = String(text)
    .replace(/[\u200B-\u200D\uFEFF\u00A0]/g, '')
    .trim()
    .toLowerCase()
    .replace(/[٠-٩]/g, (d) => String('٠١٢٣٤٥٦٧٨٩'.indexOf(d)));
  if (clean.endsWith('.0')) {
    clean = clean.substring(0, clean.length - 2);
  }
  return clean;
}

interface RepLookupModalProps {
  isOpen: boolean;
  onClose: () => void;
  machinesResults: MachineSummary[];
  expandedRows: ExpandedRow[];
  sheet1?: SheetRow[];
  sheet2?: SheetRow[];
  sheet3?: SheetRow[];
  colM1?: string;
  colM2?: string;
  colM3?: string;
  colAcc?: string;
  colRep?: string;
  colCashAcc?: string;
  colCashRep?: string;
  initialRepQuery?: string;
  initialMode?: 'rep' | 'machine';
  onFilterMainTable: (query: string) => void;
  onLoadSample?: () => void;
}

export const RepLookupModal: React.FC<RepLookupModalProps> = ({
  isOpen,
  onClose,
  machinesResults,
  expandedRows,
  sheet1 = [],
  sheet2 = [],
  sheet3 = [],
  colM1 = 'رقم الماكينة',
  colM2 = 'رقم الماكينة',
  colM3 = 'رقم الماكينة',
  colAcc = 'رقم حساب المندوب',
  colRep = 'اسم المندوب',
  colCashAcc = 'رقم حساب الكاش',
  colCashRep = 'اسم مسؤول الكاش',
  initialRepQuery = '',
  initialMode = 'rep',
  onFilterMainTable,
}) => {
  const [activeTab, setActiveTab] = useState<'rep' | 'machine'>(initialMode);
  const [lookupViewMode, setLookupViewMode] = useState<'cards' | 'table'>('table');
  const [repFilterMode, setRepFilterMode] = useState<'all' | 'duplicated' | 'cash' | 'payment' | 'both' | 'shared'>('all');
  const [repListDisplayMode, setRepListDisplayMode] = useState<'occurrences' | 'summary'>('occurrences');
  const [searchQuery, setSearchQuery] = useState(initialRepQuery || '');
  const [debouncedSearchQuery, setDebouncedSearchQuery] = useState(initialRepQuery || '');
  const [visibleMachinesLimit, setVisibleMachinesLimit] = useState(100);

  // Reset limit and filter mode when query changes
  useEffect(() => {
    setVisibleMachinesLimit(100);
    setRepFilterMode('all');
  }, [debouncedSearchQuery]);

  // Debounce search query changes
  useEffect(() => {
    const timer = setTimeout(() => {
      setDebouncedSearchQuery(searchQuery);
    }, 400); // 400ms debounce
    return () => clearTimeout(timer);
  }, [searchQuery]);

  // Set search query instantly for dropdowns, chips, links, or when modal opens
  const setSearchQueryInstant = (val: string) => {
    setSearchQuery(val);
    setDebouncedSearchQuery(val);
  };

  // Reset or set search query when modal opens
  useEffect(() => {
    if (isOpen) {
      setSearchQueryInstant(initialRepQuery || '');
      setActiveTab(initialMode || 'rep');
      setRepFilterMode('all');
    }
  }, [isOpen, initialRepQuery, initialMode]);

  // Index Sheets directly for ultra-fast lookup with ZERO duplicates & ZERO lag
  const sheet2Index = useMemo(() => {
    interface RepMachineOccurrence {
      occurrenceId: string;
      machine: string;
      cleanMachine: string;
      type: 'payment' | 'cash';
      sheetSource: 'payments' | 'cash';
      repAccount: string;
      repName: string;
      occurrenceIndex: number;
      totalForMachine: number;
      rowNumber: number;
    }

    interface RepMachineItem {
      machine: string;
      cleanMachine: string;
      cashDisplayMachine: string;
      paymentDisplayMachine: string;
      type: 'payment' | 'cash' | 'both';
      paymentCount: number;
      cashCount: number;
      totalOccurrences: number;
      occurrences: RepMachineOccurrence[];
    }

    interface RepIndexItem {
      account: string;
      name: string;
      machinesMap: Map<string, RepMachineItem>;
      allOccurrences: RepMachineOccurrence[];
    }

    const machineRepsTracker = new Map<string, Map<string, { account: string; name: string; type: 'payment' | 'cash' }>>();
    const repMap = new Map<string, RepIndexItem>();
    const repLookupMap = new Map<string, RepIndexItem>();

    if ((!sheet2 || sheet2.length === 0) && (!sheet3 || sheet3.length === 0)) {
      return {
        repList: [] as Array<{
          account: string;
          name: string;
          machineCount: number;
          duplicatedMachinesCount: number;
          totalLinksCount: number;
          normAccount: string;
          normName: string;
        }>,
        machineList: [] as Array<{
          machine: string;
          repCount: number;
          reps: Array<{ account: string; name: string; type: 'payment' | 'cash' }>;
          status: 'single' | 'multi' | 'none';
        }>,
        machineToRepsMap: new Map<string, Array<{ account: string; name: string; type: 'payment' | 'cash' }>>(),
        repLookupMap,
        repMap,
      };
    }

    // Process Sheet 2 (Payments / Donor 1 / Acceptor 1)
    if (sheet2 && sheet2.length > 0) {
      const sampleRow = sheet2[0];
      const keys = Object.keys(sampleRow);
      const mCol = colM2 && sampleRow[colM2] !== undefined
        ? colM2
        : keys.find((k) => /^(acceptor1|acceptor_1|acceptor 1|acceptor|ماكينة المدفوعات|ماكينه المدفوعات)$/i.test(k))
          || keys.find((k) => /(acceptor.*1|^acceptor$|ماكين.*مدفوعات)/i.test(k))
          || keys.find((k) => /ماكينة|جهاز|pos|machine|sn|id|acceptor/i.test(k)) || keys[0];

      const accCol = colAcc && sampleRow[colAcc] !== undefined
        ? colAcc
        : keys.find((k) => /^(doner1|donor1|doner_1|donor_1|doner 1|donor 1|doner|donor|رقم حساب المدفوعات|حساب المدفوعات)$/i.test(k))
          || keys.find((k) => /(doner.*1|donor.*1|^doner$|^donor$|حساب.*مدفوعات|رقم.*المدفوعات)/i.test(k))
          || keys.find((k) => /حساب|كود|رقم.*مندوب|account|acc|code|doner|donor/i.test(k)) || '';

      const repCol = colRep && sampleRow[colRep] !== undefined
        ? colRep
        : keys.find((k) => /اسم.*مندوب|اسم.*العميل|اسم.*المستخدم|اسم|rep.*name|agent.*name|name/i.test(k))
          || accCol || keys[1] || keys[0];

      for (let i = 0; i < sheet2.length; i++) {
        const row = sheet2[i];
        const rawM = row[mCol];
        if (rawM === undefined || rawM === null || String(rawM).trim() === '') continue;
        const rawMStr = String(rawM).trim();
        const normalizedKey = normalizeMachineId(rawMStr);
        if (!normalizedKey) continue;

        // Check if raw machine starts with 7- or ٧- (explicit Cash)
        const isCashRecord = rawMStr.startsWith('7-') || rawMStr.startsWith('٧-');
        const displayMachine = isCashRecord ? (rawMStr.startsWith('7-') || rawMStr.startsWith('٧-') ? rawMStr : `7-${normalizedKey}`) : normalizedKey;

        const rawRep = row[repCol];
        const rawAcc = accCol && row[accCol] !== undefined ? String(row[accCol]).trim() : '';

        const repName = rawRep !== undefined && rawRep !== null && String(rawRep).trim() !== ''
          ? String(rawRep).trim()
          : (rawAcc ? `مندوب (${rawAcc})` : 'لا يوجد مندوب');

        const accVal = rawAcc || 'غير متوفر';
        const normAcc = accVal !== 'غير متوفر' ? normalizeText(accVal) : '';
        const normName = repName && repName !== 'لا يوجد مندوب' ? normalizeText(repName) : '';

        const recordType: 'payment' | 'cash' = isCashRecord ? 'cash' : 'payment';

        // 1. Index by machine: track UNIQUE reps only on this machine
        let mTracker = machineRepsTracker.get(normalizedKey);
        if (!mTracker) {
          mTracker = new Map();
          machineRepsTracker.set(normalizedKey, mTracker);
        }
        const repDedupeKey = `${normAcc || 'noacc'}__${normName || 'noname'}__${recordType}`;
        if (!mTracker.has(repDedupeKey)) {
          mTracker.set(repDedupeKey, { account: accVal, name: repName, type: recordType });
        }

        // 2. Index by rep: track machines and duplicate occurrences for this rep
        const repPrimaryId = normAcc || normName;
        if (repPrimaryId) {
          let repObj = repMap.get(repPrimaryId);
          if (!repObj) {
            repObj = {
              account: accVal,
              name: repName,
              machinesMap: new Map(),
              allOccurrences: [],
            };
            repMap.set(repPrimaryId, repObj);
          }
          if (accVal !== 'غير متوفر' && (repObj.account === 'غير متوفر' || !repObj.account)) {
            repObj.account = accVal;
          }
          if (repName && !repName.startsWith('مندوب (') && !repName.startsWith('مسؤول كاش (')) {
            repObj.name = repName;
          }

          // Add or increment machine occurrence in rep's unique map
          let existingMachine = repObj.machinesMap.get(normalizedKey);
          if (!existingMachine) {
            existingMachine = {
              machine: displayMachine,
              cleanMachine: normalizedKey,
              cashDisplayMachine: `7-${normalizedKey}`,
              paymentDisplayMachine: normalizedKey,
              type: recordType,
              paymentCount: isCashRecord ? 0 : 1,
              cashCount: isCashRecord ? 1 : 0,
              totalOccurrences: 1,
              occurrences: [],
            };
            repObj.machinesMap.set(normalizedKey, existingMachine);
          } else {
            if (isCashRecord) {
              existingMachine.cashCount += 1;
            } else {
              existingMachine.paymentCount += 1;
            }
            existingMachine.totalOccurrences += 1;
            if (existingMachine.paymentCount > 0 && existingMachine.cashCount > 0) {
              existingMachine.type = 'both';
            }
          }

          const currentOccIndex = existingMachine.occurrences.length + 1;
          const occurrenceObj: RepMachineOccurrence = {
            occurrenceId: `sheet2-${i + 1}-${normalizedKey}-${currentOccIndex}`,
            machine: displayMachine,
            cleanMachine: normalizedKey,
            type: recordType,
            sheetSource: isCashRecord ? 'cash' : 'payments',
            repAccount: accVal,
            repName,
            occurrenceIndex: currentOccIndex,
            totalForMachine: existingMachine.totalOccurrences,
            rowNumber: i + 1,
          };
          existingMachine.occurrences.push(occurrenceObj);
          repObj.allOccurrences.push(occurrenceObj);

          if (normAcc && !repLookupMap.has(normAcc)) {
            repLookupMap.set(normAcc, repObj);
          }
          if (normName && !repLookupMap.has(normName)) {
            repLookupMap.set(normName, repObj);
          }
        }
      }
    }

    // Process Sheet 3 (Cash / Donor 2 / Acceptor 2)
    if (sheet3 && sheet3.length > 0) {
      const sampleRow3 = sheet3[0];
      const keys3 = Object.keys(sampleRow3);
      const mCol3 = colM3 && sampleRow3[colM3] !== undefined
        ? colM3
        : keys3.find((k) => /^(acceptor2|acceptor_2|acceptor 2|رقم ماكينه الكاش|ماكينه الكاش|ماكينة الكاش)$/i.test(k))
          || keys3.find((k) => /(acceptor.*2|ماكين.*كاش|ماكينة.*الكاش)/i.test(k))
          || keys3.find((k) => /ماكينة|جهاز|pos|machine|sn|id|acceptor/i.test(k)) || keys3[0];

      const accCol3 = colCashAcc && sampleRow3[colCashAcc] !== undefined
        ? colCashAcc
        : keys3.find((k) => /^(doner2|donor2|doner_2|donor_2|doner 2|donor 2|رقم الكاش المدفوعات|رقم الكاش|حساب الكاش)$/i.test(k))
          || keys3.find((k) => /(doner.*2|donor.*2|كاش.*مدفوعات|رقم.*الكاش|حساب.*الكاش)/i.test(k))
          || keys3.find((k) => /حساب|كود|رقم.*مندوب|account|acc|code|doner|donor/i.test(k)) || '';

      const repCol3 = colCashRep && sampleRow3[colCashRep] !== undefined
        ? colCashRep
        : keys3.find((k) => /اسم.*مندوب|اسم.*العميل|اسم.*المستخدم|اسم|rep.*name|agent.*name|name/i.test(k))
          || accCol3 || keys3[1] || keys3[0];

      for (let i = 0; i < sheet3.length; i++) {
        const row = sheet3[i];
        const rawM = row[mCol3];
        if (rawM === undefined || rawM === null || String(rawM).trim() === '') continue;
        const rawMStr = String(rawM).trim();
        const normalizedKey = normalizeMachineId(rawMStr);
        if (!normalizedKey) continue;

        // Cash machines MUST have 7- prefix
        const displayMachine = rawMStr.startsWith('7-') || rawMStr.startsWith('٧-') ? rawMStr : `7-${normalizedKey}`;

        const rawRep = row[repCol3];
        const rawAcc = accCol3 && row[accCol3] !== undefined ? String(row[accCol3]).trim() : '';

        const repName = rawRep !== undefined && rawRep !== null && String(rawRep).trim() !== ''
          ? String(rawRep).trim()
          : (rawAcc ? `مسؤول كاش (${rawAcc})` : 'لا يوجد مسؤول كاش');

        const accVal = rawAcc || 'غير متوفر';
        const normAcc = accVal !== 'غير متوفر' ? normalizeText(accVal) : '';
        const normName = repName && repName !== 'لا يوجد مسؤول كاش' ? normalizeText(repName) : '';

        // 1. Index by machine: track UNIQUE reps only on this machine
        let mTracker = machineRepsTracker.get(normalizedKey);
        if (!mTracker) {
          mTracker = new Map();
          machineRepsTracker.set(normalizedKey, mTracker);
        }
        const repDedupeKey = `${normAcc || 'noacc'}__${normName || 'noname'}__cash`;
        if (!mTracker.has(repDedupeKey)) {
          mTracker.set(repDedupeKey, { account: accVal, name: repName, type: 'cash' });
        }

        // 2. Index by rep: track machines and duplicate occurrences for this rep
        const repPrimaryId = normAcc || normName;
        if (repPrimaryId) {
          let repObj = repMap.get(repPrimaryId);
          if (!repObj) {
            repObj = {
              account: accVal,
              name: repName,
              machinesMap: new Map(),
              allOccurrences: [],
            };
            repMap.set(repPrimaryId, repObj);
          }
          if (accVal !== 'غير متوفر' && (repObj.account === 'غير متوفر' || !repObj.account)) {
            repObj.account = accVal;
          }
          if (repName && !repName.startsWith('مسؤول كاش (') && !repName.startsWith('مندوب (')) {
            repObj.name = repName;
          }

          // Add or increment machine occurrence in rep's unique map
          let existingMachine = repObj.machinesMap.get(normalizedKey);
          if (!existingMachine) {
            existingMachine = {
              machine: displayMachine,
              cleanMachine: normalizedKey,
              cashDisplayMachine: `7-${normalizedKey}`,
              paymentDisplayMachine: normalizedKey,
              type: 'cash',
              paymentCount: 0,
              cashCount: 1,
              totalOccurrences: 1,
              occurrences: [],
            };
            repObj.machinesMap.set(normalizedKey, existingMachine);
          } else {
            existingMachine.cashCount += 1;
            existingMachine.totalOccurrences += 1;
            if (existingMachine.paymentCount > 0) {
              existingMachine.type = 'both';
            }
          }

          const currentOccIndex = existingMachine.occurrences.length + 1;
          const occurrenceObj: RepMachineOccurrence = {
            occurrenceId: `sheet3-${i + 1}-${normalizedKey}-${currentOccIndex}`,
            machine: displayMachine,
            cleanMachine: normalizedKey,
            type: 'cash',
            sheetSource: 'cash',
            repAccount: accVal,
            repName,
            occurrenceIndex: currentOccIndex,
            totalForMachine: existingMachine.totalOccurrences,
            rowNumber: i + 1,
          };
          existingMachine.occurrences.push(occurrenceObj);
          repObj.allOccurrences.push(occurrenceObj);

          if (normAcc && !repLookupMap.has(normAcc)) {
            repLookupMap.set(normAcc, repObj);
          }
          if (normName && !repLookupMap.has(normName)) {
            repLookupMap.set(normName, repObj);
          }
        }
      }
    }

    // Finalize machineToRepsMap with clean unique arrays
    const machineToRepsMap = new Map<string, Array<{ account: string; name: string; type: 'payment' | 'cash' }>>();
    for (const [mId, tracker] of machineRepsTracker.entries()) {
      machineToRepsMap.set(mId, Array.from(tracker.values()));
    }

    // Update occurrences total count in all occurrences
    for (const repObj of repMap.values()) {
      for (const occ of repObj.allOccurrences) {
        const mObj = repObj.machinesMap.get(occ.cleanMachine);
        if (mObj) {
          occ.totalForMachine = mObj.totalOccurrences;
        }
      }
    }

    // Finalize distinct representatives list
    const repList = Array.from(repMap.values()).map((r) => {
      const machines = Array.from(r.machinesMap.values());
      const duplicatedMachinesCount = machines.filter((m) => m.totalOccurrences > 1).length;
      const totalLinksCount = r.allOccurrences.length || machines.reduce((sum, m) => sum + m.totalOccurrences, 0);
      return {
        account: r.account,
        name: r.name,
        machineCount: machines.length,
        duplicatedMachinesCount,
        totalLinksCount,
        machines,
        allOccurrences: r.allOccurrences,
        normAccount: normalizeText(r.account),
        normName: normalizeText(r.name),
      };
    }).sort((a, b) => b.totalLinksCount - a.totalLinksCount);

    // Finalize machineList
    const machineList = Array.from(machineToRepsMap.entries()).map(([mKey, reps]) => {
      const norm = normalizeMachineId(mKey);
      const isPureCash = reps.length > 0 && reps.every((r) => r.type === 'cash');
      const displayMachine = isPureCash
        ? (norm.startsWith('7-') || norm.startsWith('٧-') ? norm : `7-${norm}`)
        : norm;
      return {
        machine: displayMachine,
        repCount: reps.length,
        reps,
        status: reps.length === 1 ? ('single' as const) : reps.length > 1 ? ('multi' as const) : ('none' as const),
      };
    }).sort((a, b) => b.repCount - a.repCount);

    return { repList, machineList, machineToRepsMap, repLookupMap, repMap };
  }, [sheet2, sheet3, colM2, colM3, colAcc, colRep, colCashAcc, colCashRep]);

  // Unique Reps for dropdown & chips (Fast direct reference)
  const uniqueReps = useMemo(() => {
    return sheet2Index.repList;
  }, [sheet2Index]);

  // Unique Machines for dropdown & chips
  const uniqueMachines = useMemo(() => {
    if (machinesResults && machinesResults.length > 0) {
      return machinesResults;
    }
    if (sheet1.length > 0) {
      const setM = new Set<string>();
      sheet1.forEach((r) => {
        const m = String(r[colM1] || Object.values(r)[0] || '').trim();
        if (m) setM.add(m);
      });
      return Array.from(setM).map((m) => {
        const reps = sheet2Index.machineToRepsMap.get(m) || [];
        return {
          machine: m,
          repCount: reps.length,
          reps,
          status: (reps.length === 1 ? 'single' : reps.length > 1 ? 'multi' : 'none') as 'single' | 'multi' | 'none',
        };
      });
    }
    return sheet2Index.machineList;
  }, [machinesResults, sheet1, colM1, sheet2Index]);

  // Performance Optimization: Dynamic search filter with a maximum of 100 options in the DOM at once
  const filteredDropdownReps = useMemo(() => {
    if (!debouncedSearchQuery.trim()) {
      return uniqueReps.slice(0, 100); // Show top 100 by machine count
    }
    const q = debouncedSearchQuery.toLowerCase().trim();
    return uniqueReps
      .filter((r) => r.name.toLowerCase().includes(q) || r.account.toLowerCase().includes(q))
      .slice(0, 100);
  }, [uniqueReps, debouncedSearchQuery]);

  const filteredDropdownMachines = useMemo(() => {
    if (!debouncedSearchQuery.trim()) {
      return uniqueMachines.slice(0, 100);
    }
    const q = debouncedSearchQuery.toLowerCase().trim();
    return uniqueMachines
      .filter((m) => m.machine.toLowerCase().includes(q))
      .slice(0, 100);
  }, [uniqueMachines, debouncedSearchQuery]);

  // Clean chips for reps (excluding "لا يوجد مندوب" placeholders)
  const repQuickChips = useMemo(() => {
    return uniqueReps
      .filter((r) => r.name !== 'لا يوجد مندوب' || r.account !== 'غير متوفر')
      .slice(0, 8);
  }, [uniqueReps]);

  // Clean chips for machines
  const machineQuickChips = useMemo(() => {
    return uniqueMachines.slice(0, 8);
  }, [uniqueMachines]);

  // Representative Matched Data (Ultra-fast direct lookup)
  const matchedRepData = useMemo(() => {
    if (!debouncedSearchQuery.trim() || activeTab !== 'rep') {
      return null;
    }

    const q = debouncedSearchQuery.trim().toLowerCase();
    const qNorm = normalizeText(q);
    if (!qNorm) return null;

    // 1. Direct O(1) exact match
    let matchedItem = sheet2Index.repLookupMap.get(qNorm);

    // 2. If not exact hit, search in distinct representatives list
    if (!matchedItem) {
      const foundInList = sheet2Index.repList.find(
        (r) => (r.normAccount && r.normAccount.includes(qNorm)) || (r.normName && r.normName.includes(qNorm))
      );
      if (foundInList) {
        matchedItem = sheet2Index.repMap.get(foundInList.normAccount || foundInList.normName);
      }
    }

    // 3. Fallback check in expandedRows if present
    if (!matchedItem && expandedRows && expandedRows.length > 0) {
      const foundRow = expandedRows.find((r) => {
        if (r.status === 'none') return false;
        const accNorm = r.account !== 'غير متوفر' ? normalizeText(r.account) : '';
        const nameNorm = r.repName && r.repName !== 'لا يوجد مندوب' ? normalizeText(r.repName) : '';
        return (accNorm && accNorm.includes(qNorm)) || (nameNorm && nameNorm.includes(qNorm));
      });
      if (foundRow) {
        const foundAcc = normalizeText(foundRow.account);
        const foundName = normalizeText(foundRow.repName);
        matchedItem = sheet2Index.repLookupMap.get(foundAcc) || sheet2Index.repLookupMap.get(foundName);
      }
    }

    if (!matchedItem) return null;

    const primaryAccount = matchedItem.account || 'غير متوفر';
    const primaryName = matchedItem.name || `مندوب (${primaryAccount})`;
    const cleanCurrentAcc = normalizeText(primaryAccount);
    const cleanCurrentName = normalizeText(primaryName);

    // Build machines list from the pre-indexed unique machines map
    const machinesList = Array.from(matchedItem.machinesMap.values()).map((mObj) => {
      const normKey = normalizeMachineId(mObj.cleanMachine || mObj.machine);
      const allRepsOnMachine = sheet2Index.machineToRepsMap.get(normKey) || [];

      const otherReps = allRepsOnMachine.filter((r) => {
        const rAcc = normalizeText(r.account);
        const rName = normalizeText(r.name);
        if (cleanCurrentAcc && cleanCurrentAcc !== 'غير متوفر' && rAcc) {
          return rAcc !== cleanCurrentAcc;
        }
        return rName !== cleanCurrentName;
      });

      const isRepeatedForSameRep = mObj.totalOccurrences > 1;

      // Human-readable repeat details
      let repeatDescription = '';
      if (isRepeatedForSameRep) {
        if (mObj.paymentCount > 0 && mObj.cashCount > 0) {
          repeatDescription = `كاش ومدفوعات معاً (${mObj.paymentCount} مدفوعات + ${mObj.cashCount} كاش)`;
        } else if (mObj.paymentCount > 1) {
          repeatDescription = `مكررة ${mObj.paymentCount} مرات بالمدفوعات`;
        } else if (mObj.cashCount > 1) {
          repeatDescription = `مكررة ${mObj.cashCount} مرات بالكاش (7-)`;
        } else {
          repeatDescription = `مكررة ${mObj.totalOccurrences} مرات`;
        }
      }

      // Exact display label: If pure cash or starting with 7-, show 7- prefix
      const cleanId = normalizeMachineId(mObj.cleanMachine || mObj.machine);
      const displayMachine = mObj.type === 'cash'
        ? (mObj.machine.startsWith('7-') || mObj.machine.startsWith('٧-') ? mObj.machine : `7-${cleanId}`)
        : (mObj.type === 'both' ? `${cleanId} (7-${cleanId})` : cleanId);

      return {
        machine: displayMachine,
        cleanMachine: cleanId,
        type: mObj.type,
        paymentCount: mObj.paymentCount,
        cashCount: mObj.cashCount,
        sameRepOccurrences: mObj.totalOccurrences,
        isRepeatedForSameRep,
        repeatDescription,
        totalRepsOnMachine: Math.max(allRepsOnMachine.length, 1),
        repAccount: primaryAccount,
        repName: primaryName,
        otherReps,
        occurrences: mObj.occurrences || [],
        status: (allRepsOnMachine.length > 1 ? 'multi' : 'single') as 'multi' | 'single',
      };
    });

    // Build flattened all occurrences list with full metadata (all sequential rows)
    const occurrencesList = (matchedItem.allOccurrences || []).map((occ, oIdx) => {
      const normKey = normalizeMachineId(occ.cleanMachine || occ.machine);
      const allRepsOnMachine = sheet2Index.machineToRepsMap.get(normKey) || [];
      const otherReps = allRepsOnMachine.filter((r) => {
        const rAcc = normalizeText(r.account);
        const rName = normalizeText(r.name);
        if (cleanCurrentAcc && cleanCurrentAcc !== 'غير متوفر' && rAcc) {
          return rAcc !== cleanCurrentAcc;
        }
        return rName !== cleanCurrentName;
      });

      const mParent = matchedItem.machinesMap.get(normKey);
      const totalOccurrences = mParent ? mParent.totalOccurrences : occ.totalForMachine || 1;
      const isRepeatedForSameRep = totalOccurrences > 1;

      // Ensure 7- prefix for cash transactions
      const formattedMachine = occ.type === 'cash'
        ? (occ.machine.startsWith('7-') || occ.machine.startsWith('٧-') ? occ.machine : `7-${normKey}`)
        : normKey;

      const otherRepsStr = otherReps.length > 0
        ? otherReps.map((r) => `${r.name} (${r.account})`).join(' ، ')
        : 'المندوب الوحيد';

      return {
        ...occ,
        id: `occ-${oIdx + 1}-${normKey}`,
        globalIndex: oIdx + 1,
        machine: formattedMachine,
        cleanMachine: normKey,
        totalForMachine: totalOccurrences,
        isRepeatedForSameRep,
        totalRepsOnMachine: Math.max(allRepsOnMachine.length, 1),
        otherReps,
        otherRepsStr,
        repeatDescription: mParent ? mParent.type === 'both' ? 'كاش ومدفوعات' : `مكررة ${totalOccurrences} مرات` : '',
      };
    });

    const cashCount = machinesList.reduce((sum, m) => sum + (m.cashCount || (m.type === 'cash' ? 1 : 0)), 0);
    const paymentsCount = machinesList.reduce((sum, m) => sum + (m.paymentCount || (m.type === 'payment' ? 1 : 0)), 0);
    const bothCount = machinesList.filter((m) => m.type === 'both').length;
    const duplicatedForRepCount = machinesList.filter((m) => m.isRepeatedForSameRep).length;
    const totalLinksForRep = occurrencesList.length || machinesList.reduce((acc, m) => acc + m.sameRepOccurrences, 0);
    const soleCount = machinesList.filter((m) => m.totalRepsOnMachine === 1).length;
    const sharedCount = machinesList.filter((m) => m.totalRepsOnMachine > 1).length;

    return {
      repInfo: {
        account: primaryAccount,
        name: primaryName,
        totalMachines: machinesList.length,
        cashCount,
        paymentsCount,
        bothCount,
        duplicatedForRepCount,
        totalLinksForRep,
        soleCount,
        sharedCount,
      },
      machines: machinesList,
      allOccurrences: occurrencesList,
    };
  }, [debouncedSearchQuery, activeTab, expandedRows, sheet2Index]);

  // Filtered Unique Machines for current Rep view mode
  const filteredRepMachines = useMemo(() => {
    if (!matchedRepData) return [];
    const list = matchedRepData.machines;
    switch (repFilterMode) {
      case 'duplicated':
        return list.filter((m) => m.isRepeatedForSameRep);
      case 'cash':
        return list.filter((m) => 
          m.cashCount > 0 || 
          m.type === 'cash' || 
          m.type === 'both' || 
          m.machine.startsWith('7-') || 
          m.machine.startsWith('٧-') ||
          (m.repeatDescription && m.repeatDescription.toLowerCase().includes('cash')) ||
          (m.repeatDescription && m.repeatDescription.includes('كاش'))
        );
      case 'payment':
        return list.filter((m) => m.paymentCount > 0 || m.type === 'payment' || m.type === 'both');
      case 'both':
        return list.filter((m) => m.type === 'both');
      case 'shared':
        return list.filter((m) => m.totalRepsOnMachine > 1);
      default:
        return list;
    }
  }, [matchedRepData, repFilterMode]);

  // Filtered Sequential Occurrences (All Rows under each other)
  const filteredRepOccurrences = useMemo(() => {
    if (!matchedRepData) return [];
    const list = matchedRepData.allOccurrences;
    switch (repFilterMode) {
      case 'duplicated':
        return list.filter((o) => o.isRepeatedForSameRep);
      case 'cash':
        return list.filter((o) => 
          o.type === 'cash' || 
          o.sheetSource === 'cash' || 
          o.machine.startsWith('7-') || 
          o.machine.startsWith('٧-') ||
          normalizeMachineId(o.machine).startsWith('7-') ||
          (o.repeatDescription && o.repeatDescription.toLowerCase().includes('cash')) ||
          (o.repeatDescription && o.repeatDescription.includes('كاش'))
        );
      case 'payment':
        return list.filter((o) => o.type === 'payment');
      case 'both':
        return list.filter((o) => o.totalForMachine > 1);
      case 'shared':
        return list.filter((o) => o.totalRepsOnMachine > 1);
      default:
        return list;
    }
  }, [matchedRepData, repFilterMode]);

  // Machine Matched Data (الاستعلام عن ماكينة معينة)
  const matchedMachineData = useMemo(() => {
    if (!debouncedSearchQuery.trim() || activeTab !== 'machine') {
      return null;
    }

    const q = debouncedSearchQuery.trim().toLowerCase();
    const qNorm = normalizeMachineId(q);

    // Direct O(1) map lookup
    if (sheet2Index.machineToRepsMap.has(qNorm)) {
      const reps = sheet2Index.machineToRepsMap.get(qNorm)!;
      const isPureCash = reps.length > 0 && reps.every((r) => r.type === 'cash');
      const displayMachine = isPureCash ? (qNorm.startsWith('7-') ? qNorm : `7-${qNorm}`) : qNorm;
      return {
        machine: displayMachine,
        repCount: reps.length,
        status: reps.length === 1 ? 'single' : reps.length > 1 ? 'multi' : 'none' as const,
        reps: reps.map((r) => ({ account: r.account, name: r.name, type: r.type })),
      };
    }

    // CASE 1: Using machinesResults
    if (machinesResults && machinesResults.length > 0) {
      const found = machinesResults.find(
        (m) => m.machine.toLowerCase().includes(q) || normalizeMachineId(m.machine).toLowerCase().includes(qNorm)
      );
      if (found) {
        const clean = normalizeMachineId(found.machine);
        const isPureCash = found.reps.length > 0 && found.reps.every((r) => r.type === 'cash');
        const displayMachine = isPureCash ? (clean.startsWith('7-') ? clean : `7-${clean}`) : clean;
        return {
          machine: displayMachine,
          repCount: found.repCount,
          status: found.status,
          reps: found.reps,
        };
      }
    }

    // CASE 2: Fallback prefix/contains loop
    if (sheet2Index.machineToRepsMap.size > 0 || sheet1.length > 0) {
      for (const [mId, reps] of sheet2Index.machineToRepsMap.entries()) {
        const norm = normalizeMachineId(mId);
        if (mId.toLowerCase().includes(q) || norm.toLowerCase().includes(qNorm)) {
          const isPureCash = reps.length > 0 && reps.every((r) => r.type === 'cash');
          const displayMachine = isPureCash ? (norm.startsWith('7-') ? norm : `7-${norm}`) : norm;
          return {
            machine: displayMachine,
            repCount: reps.length,
            status: reps.length === 1 ? 'single' : reps.length > 1 ? 'multi' : 'none',
            reps: reps.map((r) => ({ account: r.account, name: r.name, type: r.type })),
          };
        }
      }
    }

    return null;
  }, [debouncedSearchQuery, activeTab, machinesResults, sheet2Index, sheet1]);

  if (!isOpen) return null;

  // Total records count read
  const totalSheetRecords = sheet2.length || sheet1.length || expandedRows.length;
  const isSheetLoaded = totalSheetRecords > 0;

  // Export representative statement to Excel
  const exportRepStatement = () => {
    if (!matchedRepData || (!matchedRepData.machines.length && !matchedRepData.allOccurrences.length)) return;

    const rep = matchedRepData.repInfo;

    // Sheet 1: Detailed Sequential Occurrences (All occurrences under each other with 4 rows if 4 links)
    const occurrencesData: (string | number)[][] = [
      ['كشف سجلات وارتباطات المندوب المفصلة (جميع الأسطر)'],
      ['اسم المندوب:', rep.name],
      ['رقم الحساب:', rep.account],
      ['إجمالي عدد الارتباطات والأسطر:', rep.totalLinksForRep],
      ['إجمالي عدد الماكينات الفريدة:', rep.totalMachines],
      ['ماكينات مكررة لنفس المندوب:', rep.duplicatedForRepCount],
      ['سجلات الكاش (7-):', rep.cashCount],
      ['سجلات المدفوعات:', rep.paymentsCount],
      [],
      [
        'م',
        'رقم الماكينة',
        'نوع الحساب',
        'ترتيب السجل للمندوب',
        'إجمالي تكرار الماكينة للمندوب',
        'تفاصيل نوع الارتباط',
        'حالة الملكية والتعدد مع الآخرين',
        'المناديب المشاركون على الماكينة',
        'رقم حساب المندوب',
        'اسم المندوب',
      ],
    ];

    matchedRepData.allOccurrences.forEach((occ, idx) => {
      occurrencesData.push([
        idx + 1,
        occ.machine,
        occ.type === 'cash' ? 'كاش (7-)' : 'مدفوعات',
        `سجل ${occ.occurrenceIndex} من ${occ.totalForMachine}`,
        occ.totalForMachine > 1 ? `مكررة (${occ.totalForMachine} مرات)` : 'مرة واحدة (غير مكررة)',
        occ.repeatDescription || (occ.totalForMachine > 1 ? `مكررة ${occ.totalForMachine} مرات` : 'ارتباط مفرد'),
        occ.totalRepsOnMachine === 1 ? 'خاصة بالمندوب فقط' : `مشتركة (${occ.totalRepsOnMachine} مناديب)`,
        occ.otherRepsStr,
        rep.account,
        rep.name,
      ]);
    });

    // Sheet 2: Unique Machines Aggregated Summary
    const summaryData: (string | number)[][] = [
      ['ملخص الماكينات المجمعة للمندوب'],
      ['اسم المندوب:', rep.name],
      ['رقم الحساب:', rep.account],
      ['إجمالي الماكينات الفريدة:', rep.totalMachines],
      ['إجمالي مرات الارتباط:', rep.totalLinksForRep],
      ['ماكينات مكررة لنفس المندوب:', rep.duplicatedForRepCount],
      [],
      [
        'م',
        'رقم الماكينة',
        'نوع الحساب',
        'تكرار الربط لنفس المندوب',
        'عدد مرات الارتباط',
        'تفاصيل التكرار',
        'حالة الملكية والتعدد مع الآخرين',
        'المناديب المشاركون على نفس الماكينة',
      ],
    ];

    matchedRepData.machines.forEach((m, idx) => {
      const coRepsStr = m.otherReps.length > 0
        ? m.otherReps.map((r) => `${r.name} (${r.account})`).join(' ، ')
        : 'المندوب الوحيد لهذه الماكينة';

      const stStr = m.totalRepsOnMachine === 1 ? 'خاصة بالمندوب فقط' : `مشتركة (${m.totalRepsOnMachine} مناديب)`;
      const typeStr = m.type === 'both' ? 'كاش ومدفوعات معاً' : m.type === 'cash' ? 'كاش' : 'مدفوعات';
      const repDedupeStr = m.isRepeatedForSameRep ? `مكررة (${m.sameRepOccurrences} مرات)` : 'مرة واحدة (غير مكررة)';
      const repDetailStr = m.isRepeatedForSameRep ? m.repeatDescription : 'ارتباط مفرد';

      summaryData.push([
        idx + 1,
        m.machine,
        typeStr,
        repDedupeStr,
        m.sameRepOccurrences,
        repDetailStr,
        stStr,
        coRepsStr,
      ]);
    });

    const wb = XLSX.utils.book_new();

    // 1. First Sheet: Detailed Occurrences
    const ws1 = XLSX.utils.aoa_to_sheet(occurrencesData);
    ws1['!cols'] = [
      { wch: 6 },
      { wch: 22 },
      { wch: 16 },
      { wch: 20 },
      { wch: 24 },
      { wch: 28 },
      { wch: 26 },
      { wch: 45 },
      { wch: 18 },
      { wch: 28 },
    ];
    XLSX.utils.book_append_sheet(wb, ws1, `كافة_سجلات_${rep.account}`);

    // 2. Second Sheet: Unique Summary
    const ws2 = XLSX.utils.aoa_to_sheet(summaryData);
    ws2['!cols'] = [
      { wch: 6 },
      { wch: 22 },
      { wch: 18 },
      { wch: 24 },
      { wch: 18 },
      { wch: 36 },
      { wch: 24 },
      { wch: 50 },
    ];
    XLSX.utils.book_append_sheet(wb, ws2, `ملخص_الماكينات`);

    XLSX.writeFile(wb, `كشف_ماكينات_المندوب_${rep.name}_${rep.account}.xlsx`);
  };

  // Export machine statement to Excel
  const exportMachineStatement = () => {
    if (!matchedMachineData) return;

    const m = matchedMachineData;
    const data: (string | number)[][] = [
      ['بيان المناديب المربوطين بالماكينة'],
      ['رقم الماكينة:', m.machine],
      ['حالة الماكينة:', m.status === 'single' ? 'مندوب واحد (1/1)' : m.status === 'multi' ? `متعددة المناديب (${m.repCount} مناديب)` : 'شاغرة بدون مندوب'],
      ['إجمالي المناديب المربوطين:', m.repCount],
      [],
      ['م', 'رقم حساب المندوب', 'اسم المندوب', 'ترتيب التخصيص', 'نوع التخصيص'],
    ];

    if (m.reps.length === 0) {
      data.push([1, 'غير متوفر', 'لا يوجد مندوب', '-', 'شاغرة']);
    } else {
      m.reps.forEach((r, idx) => {
        data.push([
          idx + 1,
          r.account,
          r.name,
          `مندوب ${idx + 1} من ${m.repCount}`,
          m.repCount === 1 ? 'منفرد' : 'مشترك',
        ]);
      });
    }

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [{ wch: 6 }, { wch: 20 }, { wch: 30 }, { wch: 20 }, { wch: 18 }];
    XLSX.utils.book_append_sheet(wb, ws, `ماكينة_${m.machine}`);
    XLSX.writeFile(wb, `بيان_ماكينة_${m.machine}_المناديب.xlsx`);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-[#030712]/85 backdrop-blur-md p-3 sm:p-4 transition-all font-['Cairo']">
      <div className="bg-[#0d1527] border border-[#1e293b] rounded-2xl w-full max-w-5xl shadow-[0_20px_40px_-15px_rgba(0,0,0,0.8)] overflow-hidden flex flex-col max-h-[92vh]">
        
        {/* 1. HEADER SECTION */}
        <div className="p-4 bg-[#0a101d] border-b border-slate-800/80 flex items-center justify-between gap-4">
          {/* Close button on the RIGHT in RTL */}
          <button
            onClick={onClose}
            className="w-8 h-8 rounded-lg bg-slate-800/80 hover:bg-slate-700 text-slate-400 hover:text-white flex items-center justify-center transition border border-slate-700/60 shrink-0 cursor-pointer"
            title="إغلاق النافذة"
          >
            <X className="w-4 h-4" />
          </button>

          {/* Title & Icon on the LEFT in RTL */}
          <div className="flex items-center gap-3">
            <div className="text-left sm:text-right">
              <div className="flex items-center gap-2 flex-wrap justify-end">
                <span className="text-[10px] font-mono font-bold bg-amber-500/10 text-amber-400 border border-amber-500/30 px-2.5 py-0.5 rounded-full">
                  Instant Lookup Engine
                </span>
                <h2 className="text-base font-extrabold text-white">
                  مركز الاستعلام السريع عن المناديب والماكينات
                </h2>
              </div>
              <p className="text-xs text-slate-400 font-medium mt-0.5">
                استعلم فورياً عن كافة الماكينات المربوطة بمندوب معين أو افحص ماكينة محددة لمعرفة كافة المناديب المرتبطين بها.
              </p>
            </div>
            <div className="w-10 h-10 rounded-xl bg-blue-600/20 text-blue-400 border border-blue-500/30 flex items-center justify-center text-lg shrink-0">
              {activeTab === 'rep' ? (
                <User className="w-5 h-5 text-blue-400" />
              ) : (
                <Smartphone className="w-5 h-5 text-indigo-400" />
              )}
            </div>
          </div>
        </div>

        {/* 2. MODE SELECTOR (استعلام عن مندوب vs استعلام عن ماكينة) */}
        <div className="px-4 pt-3 pb-2 bg-[#090e1a] border-b border-slate-800/60 flex items-center justify-between gap-3 flex-wrap">
          <div className="inline-flex rounded-xl p-1 bg-[#050811] border border-slate-800 text-xs font-bold shrink-0">
            <button
              onClick={() => {
                setActiveTab('rep');
                setSearchQueryInstant('');
              }}
              className={`px-4 py-2 rounded-lg flex items-center gap-2 transition-all cursor-pointer ${
                activeTab === 'rep'
                  ? 'bg-blue-600 text-white shadow-md shadow-blue-600/30'
                  : 'text-slate-400 hover:text-white'
              }`}
            >
              <User className="w-3.5 h-3.5" />
              <span>1. استعلام عن مندوب (حساب / اسم)</span>
            </button>

            <button
              onClick={() => {
                setActiveTab('machine');
                setSearchQueryInstant('');
              }}
              className={`px-4 py-2 rounded-lg flex items-center gap-2 transition-all cursor-pointer ${
                activeTab === 'machine'
                  ? 'bg-indigo-600 text-white shadow-md shadow-indigo-600/30'
                  : 'text-slate-400 hover:text-white'
              }`}
            >
              <Smartphone className="w-3.5 h-3.5" />
              <span>2. استعلام عن ماكينة (POS / SN)</span>
            </button>
          </div>

          {/* Status Capsule Indicator */}
          <div className="flex items-center gap-2 text-xs font-mono">
            {isSheetLoaded ? (
              <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-emerald-500/10 text-emerald-400 border border-emerald-500/20 font-bold">
                <CheckCircle2 className="w-3.5 h-3.5 text-emerald-400" />
                <span>تم قراءة الشيت بنجاح ({totalSheetRecords.toLocaleString()} سجل)</span>
              </span>
            ) : (
              <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-amber-500/10 text-amber-400 border border-amber-500/20 font-bold">
                <AlertCircle className="w-3.5 h-3.5 text-amber-400" />
                <span>لم يتم رفع شيت بعد</span>
              </span>
            )}
          </div>
        </div>

        {/* 3. SEARCH & CHIPS TOOLBAR */}
        <div className="p-4 bg-[#0a0f1a] border-b border-slate-800/80 space-y-3">
          {/* Input Row */}
          <div className="grid grid-cols-1 md:grid-cols-12 gap-3">
            {/* Search Input (Item 1 in RTL - Wide Input on Right) */}
            <div className="md:col-span-8 relative">
              <Search className="w-4 h-4 absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400 text-xs pointer-events-none" />
              <input
                type="text"
                autoComplete="off"
                autoCorrect="off"
                spellCheck="false"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder={
                  activeTab === 'rep'
                    ? 'ابحث برقم حساب المندوب أو اسمه (مثال: 660508 أو محمد)...'
                    : 'ابحث برقم الماكينة (مثال: POS-100 أو 20500)...'
                }
                className="w-full bg-[#050811] border border-slate-700/80 rounded-xl pr-4 pl-16 py-2.5 text-xs text-slate-100 placeholder-slate-500 focus:outline-none focus:border-blue-500 font-mono font-semibold transition shadow-inner"
              />
              {searchQuery && (
                <button
                  type="button"
                  onClick={() => setSearchQueryInstant('')}
                  className="absolute left-8 top-1/2 -translate-y-1/2 w-5 h-5 rounded-full bg-slate-800 hover:bg-slate-700 text-slate-400 hover:text-white flex items-center justify-center text-xs transition cursor-pointer"
                  title="مسح البحث"
                >
                  <X className="w-3 h-3" />
                </button>
              )}
            </div>

            {/* Dropdown Select (Item 2 in RTL - Left Dropdown) */}
            <div className="md:col-span-4">
              {activeTab === 'rep' ? (
                <select
                  value={searchQuery}
                  onChange={(e) => setSearchQueryInstant(e.target.value)}
                  className="w-full bg-[#050811] border border-slate-700/80 rounded-xl px-3 py-2.5 text-xs text-blue-400 font-bold focus:outline-none focus:border-blue-500 transition cursor-pointer shadow-inner truncate"
                >
                  <option value="">-- اختر مندوب من القائمة ({uniqueReps.length}) --</option>
                  {filteredDropdownReps.map((r, rIdx) => (
                    <option key={`${r.account}-${rIdx}`} value={r.account !== 'غير متوفر' ? r.account : r.name}>
                      {r.name} {r.account !== 'غير متوفر' ? `(حساب: ${r.account})` : ''} • [{r.machineCount} ماكينة]{r.duplicatedMachinesCount > 0 ? ` • [🔁 ${r.duplicatedMachinesCount} مكررة]` : ''}
                    </option>
                  ))}
                </select>
              ) : (
                <select
                  value={searchQuery}
                  onChange={(e) => setSearchQueryInstant(e.target.value)}
                  className="w-full bg-[#050811] border border-slate-700/80 rounded-xl px-3 py-2.5 text-xs text-indigo-400 font-bold focus:outline-none focus:border-indigo-500 transition cursor-pointer shadow-inner truncate"
                >
                  <option value="">-- اختر ماكينة من القائمة ({uniqueMachines.length}) --</option>
                  {filteredDropdownMachines.map((m, mIdx) => (
                    <option key={`${m.machine}-${mIdx}`} value={m.machine}>
                      {m.machine} • [{m.repCount} مناديب] {m.status === 'multi' ? '⚠️ مشتركة' : m.status === 'single' ? '✅ منفردة' : '❌ شاغرة'}
                    </option>
                  ))}
                </select>
              )}
            </div>
          </div>

          {/* Chips Row (أبرز المناديب أو أبرز الماكينات) */}
          <div className="flex items-center gap-2 overflow-x-auto pt-1 text-xs select-none scrollbar-thin">
            <span className="text-slate-400 text-xs shrink-0 font-bold ml-1">
              {activeTab === 'rep' ? 'أبرز المناديب:' : 'أبرز الماكينات:'}
            </span>

            {activeTab === 'rep' ? (
              repQuickChips.length > 0 ? (
                repQuickChips.map((r, rIdx) => {
                  const isActive =
                    searchQuery.trim().toLowerCase() === r.account.toLowerCase() ||
                    searchQuery.trim().toLowerCase() === r.name.toLowerCase();

                  return (
                    <button
                      key={`${r.account}-${rIdx}`}
                      onClick={() => setSearchQueryInstant(r.account !== 'غير متوفر' ? r.account : r.name)}
                      className={`px-3 py-1 rounded-lg shrink-0 transition cursor-pointer flex items-center gap-1.5 font-bold ${
                        isActive
                          ? 'bg-blue-600 text-white border border-blue-500 shadow-md shadow-blue-600/20'
                          : 'bg-slate-800/60 hover:bg-slate-800 text-slate-300 border border-slate-700/70'
                      }`}
                    >
                      <span>
                        {r.name !== 'لا يوجد مندوب' ? r.name : ''}{' '}
                        {r.account !== 'غير متوفر' ? `(${r.account})` : ''}
                      </span>
                      {r.duplicatedMachinesCount > 0 && (
                        <span className="text-[10px] px-1 py-0.2 rounded bg-purple-500/20 text-purple-300 font-mono" title={`يحتوي على ${r.duplicatedMachinesCount} ماكينات مكررة`}>
                          🔁
                        </span>
                      )}
                      {isActive && <CheckCircle2 className="w-3 h-3" />}
                    </button>
                  );
                })
              ) : (
                <span className="text-xs text-slate-500">
                  {isSheetLoaded ? 'لا توجد مناديب متاحة للاختيار.' : 'ارفع شيت المناديب أولاً لتظهر هنا.'}
                </span>
              )
            ) : (
              machineQuickChips.length > 0 ? (
                machineQuickChips.map((m, mIdx) => {
                  const isActive = searchQuery.trim().toLowerCase() === m.machine.toLowerCase();
                  return (
                    <button
                      key={`${m.machine}-${mIdx}`}
                      onClick={() => setSearchQueryInstant(m.machine)}
                      className={`px-3 py-1 rounded-lg shrink-0 transition cursor-pointer flex items-center gap-1.5 font-bold font-mono ${
                        isActive
                          ? 'bg-indigo-600 text-white border border-indigo-500 shadow-md shadow-indigo-600/20'
                          : 'bg-slate-800/60 hover:bg-slate-800 text-slate-300 border border-slate-700/70'
                      }`}
                    >
                      <span>{m.machine}</span>
                      <span className="text-[10px] px-1.5 py-0.2 rounded bg-slate-900/80 text-slate-400 font-sans">
                        {m.repCount} مناديب
                      </span>
                      {isActive && <CheckCircle2 className="w-3 h-3" />}
                    </button>
                  );
                })
              ) : (
                <span className="text-xs text-slate-500">
                  {isSheetLoaded ? 'لا توجد ماكينات متاحة للاختيار.' : 'ارفع شيت الماكينات أولاً لتظهر هنا.'}
                </span>
              )
            )}
          </div>
        </div>

        {/* 4. SCROLLABLE BODY */}
        <div className="p-4 overflow-y-auto space-y-4 bg-[#080c14] flex-1">
          
          {/* STATE A: NO SEARCH QUERY YET (Clean State confirming sheet is read) */}
          {!searchQuery.trim() && (
            <div className="py-12 px-6 text-center space-y-4 bg-[#090e1a] rounded-2xl border border-slate-800/80 shadow-inner">
              <div className="w-16 h-16 rounded-3xl bg-blue-600/10 border border-blue-500/20 flex items-center justify-center mx-auto text-blue-400">
                <FileCheck className="w-8 h-8 text-blue-400" />
              </div>
              
              <div className="space-y-1 max-w-md mx-auto">
                <h3 className="text-base font-extrabold text-white">
                  {isSheetLoaded ? 'الشيت مقروء بالكامل وجاهز للاستعلام' : 'في انتظار رفع الشيتات'}
                </h3>
                <p className="text-xs text-slate-400 leading-relaxed">
                  {isSheetLoaded
                    ? `تم تحليل ${totalSheetRecords.toLocaleString()} سجلاً بنجاح. اكتب رقم الحساب أو الماكينة في شريط البحث أعلاه أو اضغط على أحد المناديب/الماكينات لعرض التقرير الفوري.`
                    : 'يرجى رفع شيت الماكينات وشيت المناديب من الشاشة الرئيسية، أو اختيار «عينة فورية» للاختبار.'}
                </p>
              </div>

              {/* Quick statistics summary cards */}
              {isSheetLoaded && (
                <div className="grid grid-cols-2 sm:grid-cols-3 gap-3 max-w-lg mx-auto pt-2">
                  <div className="p-3 bg-[#050811] rounded-xl border border-slate-800 text-center">
                    <span className="text-[10px] text-slate-400 font-bold block">إجمالي السجلات</span>
                    <span className="text-lg font-mono font-black text-blue-400">{totalSheetRecords.toLocaleString()}</span>
                  </div>
                  <div className="p-3 bg-[#050811] rounded-xl border border-slate-800 text-center">
                    <span className="text-[10px] text-slate-400 font-bold block">المناديب الفريدة</span>
                    <span className="text-lg font-mono font-black text-emerald-400">{uniqueReps.length.toLocaleString()}</span>
                  </div>
                  <div className="p-3 bg-[#050811] rounded-xl border border-slate-800 text-center col-span-2 sm:col-span-1">
                    <span className="text-[10px] text-slate-400 font-bold block">الماكينات المفحوصة</span>
                    <span className="text-lg font-mono font-black text-indigo-400">{uniqueMachines.length.toLocaleString()}</span>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* SEARCHING LOADER */}
          {searchQuery.trim() && searchQuery !== debouncedSearchQuery && (
            <div className="py-16 text-center space-y-4 bg-[#090e1a] rounded-2xl border border-slate-800/80 shadow-inner flex flex-col items-center justify-center animate-in fade-in duration-150">
              <div className="w-10 h-10 rounded-full border-4 border-blue-500/20 border-t-blue-500 animate-spin"></div>
              <div className="space-y-1">
                <h4 className="text-sm font-extrabold text-blue-400">جاري البحث والتحليل فورياً...</h4>
                <p className="text-[11px] text-slate-400">نعمل على فحص وتدقيق السجلات لتوفير أدق المعلومات بدون تهنيج.</p>
              </div>
            </div>
          )}

          {/* STATE B: REP LOOKUP RESULTS */}
          {searchQuery.trim() && searchQuery === debouncedSearchQuery && activeTab === 'rep' && (
            matchedRepData ? (
              <div className="space-y-4 animate-in fade-in duration-200">
                {/* Rep Profile Card */}
                <div className="p-4 rounded-2xl bg-[#0b1222] border border-slate-800 shadow-md">
                  <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
                    {/* Rep Info */}
                    <div className="flex items-center gap-3.5">
                      <div className="w-12 h-12 rounded-2xl bg-blue-600/20 border border-blue-500/30 text-blue-400 flex items-center justify-center text-xl font-bold">
                        <User className="w-6 h-6" />
                      </div>
                      <div>
                        <div className="flex items-center gap-2 flex-wrap">
                          <h3 className="text-base font-extrabold text-white">
                            {matchedRepData.repInfo.name}
                          </h3>
                          <span className="font-mono text-xs font-bold px-2.5 py-0.5 rounded-lg bg-blue-950 text-blue-400 border border-blue-500/30">
                            حساب: {matchedRepData.repInfo.account}
                          </span>
                          {matchedRepData.repInfo.duplicatedForRepCount > 0 && (
                            <span className="inline-flex items-center gap-1 font-mono text-[11px] font-black px-2.5 py-0.5 rounded-lg bg-purple-950/80 text-purple-300 border border-purple-500/40 shadow-sm animate-pulse">
                              <Repeat className="w-3.5 h-3.5" />
                              <span>{matchedRepData.repInfo.duplicatedForRepCount} ماكينات مكررة</span>
                            </span>
                          )}
                        </div>
                        <p className="text-xs text-slate-400 mt-1">
                          بيان تفصيلي بجميع ماكينات هذا المندوب وحالة الارتباط والتكرار لكل جهاز.
                        </p>
                      </div>
                    </div>

                    {/* KPI Pills */}
                    <div className="flex items-center gap-2 flex-wrap w-full md:w-auto justify-end">
                      {/* Total Machines & Total Links KPI Button */}
                      <button
                        type="button"
                        onClick={() => {
                          setRepFilterMode('all');
                        }}
                        className={`px-3.5 py-2 rounded-xl text-center min-w-[95px] transition cursor-pointer border ${
                          repFilterMode === 'all'
                            ? 'bg-blue-600/20 border-blue-500 text-white shadow-md'
                            : 'bg-[#050811] border-slate-800 text-slate-300 hover:border-slate-700'
                        }`}
                        title="عرض إجمالي كافة الماكينات والارتباطات"
                      >
                        <span className="text-[10px] text-slate-400 font-bold block">إجمالي الماكينات</span>
                        <div className="flex items-baseline justify-center gap-1">
                          <span className="text-base font-mono font-black text-blue-400">
                            {matchedRepData.repInfo.totalLinksForRep}
                          </span>
                          <span className="text-[10px] text-slate-400 font-bold">
                            ({matchedRepData.repInfo.totalMachines} فريدة)
                          </span>
                        </div>
                      </button>

                      {/* Rep Duplicated Machines KPI Button */}
                      <button
                        type="button"
                        onClick={() => setRepFilterMode(repFilterMode === 'duplicated' ? 'all' : 'duplicated')}
                        className={`px-3.5 py-2 rounded-xl text-center min-w-[95px] transition cursor-pointer border ${
                          matchedRepData.repInfo.duplicatedForRepCount > 0
                            ? repFilterMode === 'duplicated'
                              ? 'bg-purple-600 text-white border-purple-400 shadow-lg shadow-purple-900/40'
                              : 'bg-purple-950/40 border-purple-500/50 hover:border-purple-400 text-purple-300 shadow-sm'
                            : 'bg-[#050811] border-slate-800 text-slate-400'
                        }`}
                        title="اضغط لتصفية الماكينات المكررة لنفس المندوب"
                      >
                        <span className={`text-[10px] font-bold flex items-center justify-center gap-1 ${
                          matchedRepData.repInfo.duplicatedForRepCount > 0 ? 'text-purple-300' : 'text-slate-400'
                        }`}>
                          <Repeat className="w-3 h-3" />
                          <span>مكررة للمندوب</span>
                        </span>
                        <span className={`text-base font-mono font-black ${
                          matchedRepData.repInfo.duplicatedForRepCount > 0 ? 'text-purple-200' : 'text-slate-400'
                        }`}>
                          {matchedRepData.repInfo.duplicatedForRepCount}
                        </span>
                      </button>

                      {/* Cash Machines Count Box */}
                      <button
                        type="button"
                        onClick={() => setRepFilterMode(repFilterMode === 'cash' ? 'all' : 'cash')}
                        className={`px-3.5 py-2 rounded-xl text-center min-w-[90px] transition cursor-pointer border ${
                          repFilterMode === 'cash'
                            ? 'bg-emerald-600 text-white border-emerald-400 shadow-md'
                            : 'bg-[#050811] border-emerald-500/40 hover:border-emerald-500 text-emerald-400'
                        }`}
                        title="اضغط لعرض ماكينات الكاش التي تبدأ بـ 7-"
                      >
                        <span className="text-[10px] text-emerald-400 font-bold flex items-center justify-center gap-1">
                          <span>💵 كاش (7-)</span>
                        </span>
                        <span className="text-base font-mono font-black text-emerald-400">
                          {matchedRepData.repInfo.cashCount}
                        </span>
                      </button>

                      {/* Payments Machines Count Box */}
                      <button
                        type="button"
                        onClick={() => setRepFilterMode(repFilterMode === 'payment' ? 'all' : 'payment')}
                        className={`px-3.5 py-2 rounded-xl text-center min-w-[90px] transition cursor-pointer border ${
                          repFilterMode === 'payment'
                            ? 'bg-blue-600 text-white border-blue-400 shadow-md'
                            : 'bg-[#050811] border-blue-500/40 hover:border-blue-500 text-blue-400'
                        }`}
                        title="اضغط لعرض ماكينات المدفوعات"
                      >
                        <span className="text-[10px] text-blue-400 font-bold flex items-center justify-center gap-1">
                          <span>💳 مدفوعات</span>
                        </span>
                        <span className="text-base font-mono font-black text-blue-400">
                          {matchedRepData.repInfo.paymentsCount}
                        </span>
                      </button>

                      {/* Cash + Payments Both Count Box (if exists) */}
                      {matchedRepData.repInfo.bothCount > 0 && (
                        <button
                          type="button"
                          onClick={() => setRepFilterMode(repFilterMode === 'both' ? 'all' : 'both')}
                          className={`px-3.5 py-2 rounded-xl text-center min-w-[90px] transition cursor-pointer border ${
                            repFilterMode === 'both'
                              ? 'bg-amber-600 text-white border-amber-400 shadow-md'
                              : 'bg-[#050811] border-amber-500/40 hover:border-amber-500 text-amber-400'
                          }`}
                        >
                          <span className="text-[10px] text-amber-400 font-bold flex items-center justify-center gap-1">
                            <span>⭐ كاش + دفع</span>
                          </span>
                          <span className="text-base font-mono font-black text-amber-400">
                            {matchedRepData.repInfo.bothCount}
                          </span>
                        </button>
                      )}

                      <button
                        type="button"
                        onClick={() => setRepFilterMode(repFilterMode === 'shared' ? 'all' : 'shared')}
                        className={`px-3.5 py-2 rounded-xl text-center min-w-[80px] transition cursor-pointer border ${
                          repFilterMode === 'shared'
                            ? 'bg-amber-600 text-white border-amber-400 shadow-md'
                            : 'bg-[#050811] border-amber-900/40 hover:border-amber-700 text-amber-400'
                        }`}
                      >
                        <span className="text-[10px] text-amber-400 font-bold block">مشتركة</span>
                        <span className="text-base font-mono font-black text-amber-400">
                          {matchedRepData.repInfo.sharedCount}
                        </span>
                      </button>
                    </div>
                  </div>
                </div>

                {/* DISTINCTIVE ALERT BANNER IF DUPLICATE MACHINES FOUND FOR THIS REP */}
                {matchedRepData.repInfo.duplicatedForRepCount > 0 && (
                  <div className="p-3.5 rounded-2xl bg-gradient-to-r from-purple-950/70 via-indigo-950/40 to-[#0b1222] border border-purple-500/40 shadow-lg flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3">
                    <div className="flex items-center gap-3">
                      <div className="w-10 h-10 rounded-xl bg-purple-600/20 text-purple-400 border border-purple-500/40 flex items-center justify-center shrink-0 shadow-inner">
                        <Repeat className="w-5 h-5 text-purple-300 animate-pulse" />
                      </div>
                      <div>
                        <div className="flex items-center gap-2 flex-wrap">
                          <span className="text-xs font-black text-purple-200">
                            ⚠️ تنبيه: تم رصد ماكينات مربوطة بنفس المندوب أكثر من مرة!
                          </span>
                          <span className="px-2.5 py-0.5 rounded-full text-[10px] font-mono font-black bg-purple-500/20 text-purple-300 border border-purple-500/30">
                            {matchedRepData.repInfo.duplicatedForRepCount} ماكينة مكررة ({matchedRepData.repInfo.totalLinksForRep} إجمالي ارتباط)
                          </span>
                        </div>
                        <p className="text-[11px] text-slate-300 mt-0.5">
                          تظهر هذه الماكينات بعلامة مميزة <span className="text-purple-300 font-bold">🔁</span> ومظللة باللون البنفسجي، ومسندة لهذا المندوب أكثر من مرة (بين الكاش والمدفوعات أو تكرار داخل نفس الشيت).
                        </p>
                      </div>
                    </div>

                    <button
                      type="button"
                      onClick={() => setRepFilterMode(repFilterMode === 'duplicated' ? 'all' : 'duplicated')}
                      className={`px-3.5 py-2 rounded-xl text-xs font-bold transition cursor-pointer flex items-center gap-1.5 shrink-0 shadow-md ${
                        repFilterMode === 'duplicated'
                          ? 'bg-purple-600 hover:bg-purple-500 text-white shadow-purple-600/30'
                          : 'bg-purple-950 hover:bg-purple-900 text-purple-200 border border-purple-500/40'
                      }`}
                    >
                      <Repeat className="w-3.5 h-3.5" />
                      <span>{repFilterMode === 'duplicated' ? 'عرض كافة الماكينات' : 'تصفية الماكينات المكررة فقط 🔁'}</span>
                    </button>
                  </div>
                )}

                {/* Machines Section for this Rep */}
                <div className="space-y-3">
                  <div className="p-3 bg-[#060a14] border border-slate-800 rounded-2xl flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                    <div className="flex items-center gap-2 flex-wrap">
                      <div className="flex items-center gap-2">
                        <Layers className="w-4 h-4 text-blue-400" />
                        <span className="text-xs font-bold text-white">
                          {repListDisplayMode === 'occurrences'
                            ? `الأسطر والارتباطات المعروضة (${filteredRepOccurrences.length} من ${matchedRepData.allOccurrences.length})`
                            : `الماكينات المعروضة (${filteredRepMachines.length} من ${matchedRepData.machines.length})`}
                        </span>
                      </div>

                      {/* Active Filter Badge */}
                      {repFilterMode !== 'all' && (
                        <div className="flex items-center gap-1.5">
                          <span className="text-[10px] font-bold px-2 py-0.5 rounded-lg bg-blue-500/10 text-blue-400 border border-blue-500/20">
                            تصفية: {
                              repFilterMode === 'duplicated' ? '🔁 المكررة للمندوب' :
                              repFilterMode === 'cash' ? '💵 الكاش فقط (7-)' :
                              repFilterMode === 'payment' ? '💳 المدفوعات فقط' :
                              repFilterMode === 'both' ? '⭐ كاش ومدفوعات معاً' :
                              '👥 المشتركة مع آخرين'
                            }
                          </span>
                          <button
                            type="button"
                            onClick={() => setRepFilterMode('all')}
                            className="text-[10px] text-slate-400 hover:text-white underline cursor-pointer"
                          >
                            (إلغاء التصفية)
                          </button>
                        </div>
                      )}
                    </div>

                    <div className="flex items-center gap-2.5 flex-wrap">
                      {/* Representation Format Toggle: Sequential Rows (Occurrences) vs Grouped Summary */}
                      <div className="inline-flex rounded-xl p-1 bg-[#050811] border border-slate-800 text-xs font-bold shrink-0">
                        <button
                          onClick={() => setRepListDisplayMode('occurrences')}
                          className={`px-3 py-1 rounded-lg flex items-center gap-1.5 transition-all cursor-pointer ${
                            repListDisplayMode === 'occurrences'
                              ? 'bg-indigo-600 text-white shadow-sm'
                              : 'text-slate-400 hover:text-white'
                          }`}
                          title="عرض كافة الأسطر تحت بعضها حتى لو تكررت الماكينة 4 مرات كما في الإكسل"
                        >
                          <span>📋 كافة الأسطر بالتفصيل ({matchedRepData.allOccurrences.length})</span>
                        </button>
                        <button
                          onClick={() => setRepListDisplayMode('summary')}
                          className={`px-3 py-1 rounded-lg flex items-center gap-1.5 transition-all cursor-pointer ${
                            repListDisplayMode === 'summary'
                              ? 'bg-blue-600 text-white shadow-sm'
                              : 'text-slate-400 hover:text-white'
                          }`}
                          title="عرض ملخص الماكينات المجمعة الفريدة"
                        >
                          <span>📦 ملخص الأجهزة الفريدة ({matchedRepData.machines.length})</span>
                        </button>
                      </div>

                      {/* View Switcher Capsule (Table vs Cards) */}
                      <div className="inline-flex rounded-xl p-1 bg-[#050811] border border-slate-800 text-xs font-bold shrink-0">
                        <button
                          onClick={() => setLookupViewMode('cards')}
                          className={`px-2.5 py-1 rounded-lg flex items-center gap-1 transition-all cursor-pointer ${
                            lookupViewMode === 'cards'
                              ? 'bg-slate-700 text-white shadow-sm'
                              : 'text-slate-400 hover:text-white'
                          }`}
                          title="عرض البطاقات والبوكسات"
                        >
                          <Layers className="w-3.5 h-3.5" />
                          <span>بطاقات</span>
                        </button>
                        <button
                          onClick={() => setLookupViewMode('table')}
                          className={`px-2.5 py-1 rounded-lg flex items-center gap-1 transition-all cursor-pointer ${
                            lookupViewMode === 'table'
                              ? 'bg-slate-700 text-white shadow-sm'
                              : 'text-slate-400 hover:text-white'
                          }`}
                          title="عرض جدول البيانات التفصيلي"
                        >
                          <TableProperties className="w-3.5 h-3.5" />
                          <span>جدول</span>
                        </button>
                      </div>

                      {/* Excel Export Button */}
                      <button
                        onClick={exportRepStatement}
                        className="px-3 py-1.5 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold flex items-center gap-1.5 transition cursor-pointer shadow-md"
                        title="تصدير شيت الإكسل الشامل (الأسطر المكررة + ملخص الأجهزة)"
                      >
                        <FileSpreadsheet className="w-3.5 h-3.5" />
                        <span>تصدير بيان المندوب Excel</span>
                      </button>
                    </div>
                  </div>

                  {/* View Mode Rendering: TABLE VIEW or CARDS VIEW */}
                  {lookupViewMode === 'table' ? (
                    <div className="overflow-x-auto max-h-[45vh] border border-slate-800 rounded-2xl bg-[#0b1222] shadow-xl">
                      <table className="w-full text-right text-xs border-collapse">
                        <thead className="bg-[#050811] text-slate-400 font-bold sticky top-0 border-b border-slate-800 z-10">
                          <tr>
                            <th className="p-3 w-12 text-center font-mono">م</th>
                            <th className="p-3 w-36 font-mono">رقم الماكينة</th>
                            <th className="p-3 w-32 text-center">نوع الحساب</th>
                            {repListDisplayMode === 'occurrences' ? (
                              <>
                                <th className="p-3 w-40 text-center">ترتيب السجل للمندوب</th>
                                <th className="p-3 w-40 text-center">إجمالي تكرار الماكينة</th>
                              </>
                            ) : (
                              <th className="p-3 w-48 text-center">تكرار الربط للمندوب</th>
                            )}
                            <th className="p-3 w-36 text-center">حالة الملكية والتعدد</th>
                            <th className="p-3">المناديب المشاركون على الماكينة</th>
                            <th className="p-3 w-24 text-center">إجراء</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-800/80 font-mono text-slate-200">
                          {repListDisplayMode === 'occurrences' ? (
                            /* ALL OCCURRENCES SEQUENTIAL ROWS (Matches User Screenshot) */
                            filteredRepOccurrences.slice(0, visibleMachinesLimit).map((occ, idx) => (
                              <tr
                                key={occ.id || `${occ.machine}-${idx}`}
                                className={`transition-colors ${
                                  occ.isRepeatedForSameRep
                                    ? 'bg-purple-950/20 hover:bg-purple-950/35 border-r-4 border-r-purple-500'
                                    : 'hover:bg-slate-800/50'
                                }`}
                              >
                                <td className="p-3 text-center text-slate-500 font-bold">{idx + 1}</td>
                                <td className="p-3 font-bold text-sm tracking-wide">
                                  <div className="flex items-center gap-1.5">
                                    <span className={occ.type === 'cash' ? 'text-emerald-400 font-black' : 'text-white'}>
                                      {occ.machine}
                                    </span>
                                    {occ.isRepeatedForSameRep && (
                                      <span className="text-[10px] px-1 py-0.2 rounded bg-purple-500/20 text-purple-300 font-mono" title={`سجل ${occ.occurrenceIndex} من ${occ.totalForMachine}`}>
                                        🔁 {occ.occurrenceIndex}/{occ.totalForMachine}
                                      </span>
                                    )}
                                  </div>
                                </td>
                                <td className="p-3 text-center font-bold">
                                  {occ.type === 'cash' ? (
                                    <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[10px] font-black bg-emerald-500/10 text-emerald-400 border border-emerald-500/20">
                                      💵 كاش (7-)
                                    </span>
                                  ) : (
                                    <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[10px] font-black bg-blue-500/10 text-blue-400 border border-blue-500/20">
                                      💳 مدفوعات
                                    </span>
                                  )}
                                </td>
                                <td className="p-3 text-center">
                                  <span className={`inline-block px-2.5 py-0.5 rounded-lg text-[10px] font-black font-sans border ${
                                    occ.isRepeatedForSameRep
                                      ? 'bg-purple-950/80 text-purple-300 border-purple-500/40'
                                      : 'bg-slate-800 text-slate-400 border-slate-700'
                                  }`}>
                                    سجل {occ.occurrenceIndex} من {occ.totalForMachine}
                                  </span>
                                </td>
                                <td className="p-3 text-center">
                                  {occ.isRepeatedForSameRep ? (
                                    <span className="text-purple-300 text-[10px] font-black font-sans">
                                      مكررة ({occ.totalForMachine} مرات)
                                    </span>
                                  ) : (
                                    <span className="text-slate-400 text-[10px] font-sans">
                                      مرة واحدة
                                    </span>
                                  )}
                                </td>
                                <td className="p-3 text-center">
                                  <span
                                    className={`inline-block px-2 py-0.5 rounded-lg text-[10px] font-black border ${
                                      occ.totalRepsOnMachine === 1
                                        ? 'bg-emerald-500/15 text-emerald-400 border-emerald-500/30'
                                        : 'bg-amber-500/15 text-amber-400 border-amber-500/30'
                                    }`}
                                  >
                                    {occ.totalRepsOnMachine === 1 ? 'خاصة فقط' : `مشتركة (${occ.totalRepsOnMachine} مناديب)`}
                                  </span>
                                </td>
                                <td className="p-3 font-sans">
                                  {occ.otherReps.length > 0 ? (
                                    <div className="flex items-center gap-1.5 flex-wrap">
                                      {occ.otherReps.map((r, rIdx) => (
                                        <span
                                          key={rIdx}
                                          onClick={() => setSearchQueryInstant(r.account !== 'غير متوفر' ? r.account : r.name)}
                                          className="px-2 py-0.5 rounded-lg bg-[#050811] hover:bg-slate-800 text-slate-300 border border-slate-800 hover:border-blue-500 text-[10px] hover:text-blue-400 cursor-pointer transition font-bold"
                                          title="اضغط للاستعلام عن هذا المندوب"
                                        >
                                          {r.name} {r.account !== 'غير متوفر' ? `(${r.account})` : ''} {r.type === 'cash' ? '💵' : '💳'}
                                        </span>
                                      ))}
                                    </div>
                                  ) : occ.isRepeatedForSameRep ? (
                                    <span className="text-purple-300 text-[11px] font-bold">
                                      مسندة لنفس المندوب ({occ.totalForMachine}) مرات متكررة بدون شركاء آخرين
                                    </span>
                                  ) : (
                                    <span className="text-slate-500 text-[11px]">المندوب هو المالك الوحيد ولا يوجد شركاء</span>
                                  )}
                                </td>
                                <td className="p-3 text-center">
                                  <button
                                    onClick={() => {
                                      setActiveTab('machine');
                                      setSearchQueryInstant(occ.machine);
                                    }}
                                    className="px-2.5 py-1 rounded-lg bg-indigo-600 hover:bg-indigo-500 text-white text-[10px] font-black transition cursor-pointer"
                                  >
                                    فحص &larr;
                                  </button>
                                </td>
                              </tr>
                            ))
                          ) : (
                            /* UNIQUE MACHINES SUMMARY ROWS */
                            filteredRepMachines.slice(0, visibleMachinesLimit).map((m, idx) => (
                              <tr
                                key={`${m.machine}-${idx}`}
                                className={`transition-colors ${
                                  m.isRepeatedForSameRep
                                    ? 'bg-purple-950/20 hover:bg-purple-950/35 border-r-4 border-r-purple-500'
                                    : 'hover:bg-slate-800/50'
                                }`}
                              >
                                <td className="p-3 text-center text-slate-500 font-bold">{idx + 1}</td>
                                <td className="p-3 font-bold text-white text-sm tracking-wide">
                                  <div className="flex items-center gap-1.5">
                                    <span className={m.type === 'cash' ? 'text-emerald-400 font-black' : 'text-white'}>
                                      {m.machine}
                                    </span>
                                    {m.isRepeatedForSameRep && (
                                      <span className="text-[10px] px-1 py-0.2 rounded bg-purple-500/20 text-purple-300 font-mono" title="ماكينة مكررة لنفس المندوب">
                                        🔁
                                      </span>
                                    )}
                                  </div>
                                </td>
                                <td className="p-3 text-center font-bold">
                                  {m.type === 'both' ? (
                                    <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[10px] font-black bg-amber-500/15 text-amber-300 border border-amber-500/30">
                                      ⭐ كاش + دفع ({m.cashCount}💵+{m.paymentCount}💳)
                                    </span>
                                  ) : m.type === 'cash' ? (
                                    <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[10px] font-black bg-emerald-500/10 text-emerald-400 border border-emerald-500/20">
                                      💵 كاش (7-) {m.cashCount > 1 ? `(${m.cashCount}×)` : ''}
                                    </span>
                                  ) : (
                                    <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[10px] font-black bg-blue-500/10 text-blue-400 border border-blue-500/20">
                                      💳 مدفوعات {m.paymentCount > 1 ? `(${m.paymentCount}×)` : ''}
                                    </span>
                                  )}
                                </td>
                                {/* Repetition for Same Rep Column */}
                                <td className="p-3 text-center">
                                  {m.isRepeatedForSameRep ? (
                                    <div className="flex flex-col items-center gap-1">
                                      <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-lg text-[10px] font-black bg-purple-500/20 text-purple-300 border border-purple-500/40 shadow-sm animate-pulse">
                                        <Repeat className="w-3 h-3" />
                                        <span>مكررة للمندوب ({m.sameRepOccurrences} مرات)</span>
                                      </span>
                                      <span className="text-[9px] text-purple-300 font-sans font-medium">
                                        {m.repeatDescription}
                                      </span>
                                    </div>
                                  ) : (
                                    <span className="inline-block px-2 py-0.5 rounded-lg text-[10px] font-bold bg-slate-800 text-slate-400 border border-slate-700">
                                      مرة واحدة (منفرد)
                                    </span>
                                  )}
                                </td>
                                <td className="p-3 text-center">
                                  <span
                                    className={`inline-block px-2 py-0.5 rounded-lg text-[10px] font-black border ${
                                      m.totalRepsOnMachine === 1
                                        ? 'bg-emerald-500/15 text-emerald-400 border-emerald-500/30'
                                        : 'bg-amber-500/15 text-amber-400 border-amber-500/30'
                                    }`}
                                  >
                                    {m.totalRepsOnMachine === 1 ? 'خاصة فقط' : `مشتركة (${m.totalRepsOnMachine} مناديب)`}
                                  </span>
                                </td>
                                <td className="p-3 font-sans">
                                  {m.otherReps.length > 0 ? (
                                    <div className="flex items-center gap-1.5 flex-wrap">
                                      {m.otherReps.map((r, rIdx) => (
                                        <span
                                          key={rIdx}
                                          onClick={() => setSearchQueryInstant(r.account !== 'غير متوفر' ? r.account : r.name)}
                                          className="px-2 py-0.5 rounded-lg bg-[#050811] hover:bg-slate-800 text-slate-300 border border-slate-800 hover:border-blue-500 text-[10px] hover:text-blue-400 cursor-pointer transition font-bold"
                                          title="اضغط للاستعلام عن هذا المندوب"
                                        >
                                          {r.name} {r.account !== 'غير متوفر' ? `(${r.account})` : ''} {r.type === 'cash' ? '💵' : '💳'}
                                        </span>
                                      ))}
                                    </div>
                                  ) : m.isRepeatedForSameRep ? (
                                    <span className="text-purple-300 text-[11px] font-bold">
                                      الماكينة مسندة لنفس المندوب ({m.sameRepOccurrences}) مرات متكررة بدون شركاء آخرين
                                    </span>
                                  ) : (
                                    <span className="text-slate-500 text-[11px]">المندوب هو المالك الوحيد ولا يوجد شركاء</span>
                                  )}
                                </td>
                                <td className="p-3 text-center">
                                  <button
                                    onClick={() => {
                                      setActiveTab('machine');
                                      setSearchQueryInstant(m.machine);
                                    }}
                                    className="px-2.5 py-1 rounded-lg bg-indigo-600 hover:bg-indigo-500 text-white text-[10px] font-black transition cursor-pointer"
                                  >
                                    فحص &larr;
                                  </button>
                                </td>
                              </tr>
                            ))
                          )}
                        </tbody>
                      </table>
                      {(repListDisplayMode === 'occurrences' ? filteredRepOccurrences.length : filteredRepMachines.length) > visibleMachinesLimit && (
                        <div className="p-3 text-center bg-[#050811] border-t border-slate-800">
                          <button
                            type="button"
                            onClick={() => setVisibleMachinesLimit((prev) => prev + 100)}
                            className="px-4 py-2 rounded-xl bg-blue-600/20 hover:bg-blue-600 text-blue-400 hover:text-white border border-blue-500/30 text-xs font-bold transition cursor-pointer"
                          >
                            عرض المزيد من السجلات (
                            {(repListDisplayMode === 'occurrences' ? filteredRepOccurrences.length : filteredRepMachines.length) - visibleMachinesLimit} متبقية من إجمالي{' '}
                            {repListDisplayMode === 'occurrences' ? filteredRepOccurrences.length : filteredRepMachines.length}
                            )
                          </button>
                        </div>
                      )}
                    </div>
                  ) : (
                    /* Grid of Cards */
                    <div className="overflow-y-auto max-h-[45vh] pr-1">
                      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
                        {repListDisplayMode === 'occurrences' ? (
                          /* Occurrences Cards */
                          filteredRepOccurrences.slice(0, visibleMachinesLimit).map((occ, idx) => (
                            <div
                              key={occ.id || `${occ.machine}-${idx}`}
                              className={`border rounded-2xl p-4 flex flex-col justify-between gap-3 shadow-md transition-all ${
                                occ.isRepeatedForSameRep
                                  ? 'bg-gradient-to-b from-[#18132e] to-[#0b1222] border-purple-500/50 shadow-purple-950/40 hover:border-purple-400'
                                  : 'bg-[#0b1222] border-slate-800 hover:border-blue-500/50'
                              }`}
                            >
                              {/* Card Header */}
                              <div className="flex items-center justify-between border-b border-slate-800 pb-2">
                                <div className="flex items-center gap-2">
                                  <span className="text-slate-500 font-mono text-[11px] font-bold">#{idx + 1}</span>
                                  <span className={`font-mono font-black text-sm tracking-wide ${occ.type === 'cash' ? 'text-emerald-400' : 'text-white'}`}>
                                    {occ.machine}
                                  </span>
                                </div>
                                <div className="flex flex-col items-end gap-1">
                                  <span
                                    className={`px-2 py-0.5 rounded-lg text-[10px] font-black border ${
                                      occ.totalRepsOnMachine === 1
                                        ? 'bg-emerald-500/15 text-emerald-400 border-emerald-500/30'
                                        : 'bg-amber-500/15 text-amber-400 border-amber-500/30'
                                    }`}
                                  >
                                    {occ.totalRepsOnMachine === 1 ? 'خاصة بالمندوب' : `مشتركة (${occ.totalRepsOnMachine} مناديب)`}
                                  </span>
                                </div>
                              </div>

                              {/* Repetition and Type Badges */}
                              <div className="flex items-center gap-1.5 flex-wrap">
                                {occ.isRepeatedForSameRep ? (
                                  <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-lg text-[10px] font-black bg-purple-500/20 text-purple-300 border border-purple-500/40 shadow-sm animate-pulse">
                                    <Repeat className="w-3 h-3" />
                                    <span>سجل {occ.occurrenceIndex} من {occ.totalForMachine}</span>
                                  </span>
                                ) : (
                                  <span className="px-2 py-0.5 rounded-md text-[10px] font-bold bg-slate-800 text-slate-400 border border-slate-700">
                                    سجل مفرد
                                  </span>
                                )}

                                {occ.type === 'cash' ? (
                                  <span className="px-2 py-0.5 rounded-md text-[10px] font-black bg-emerald-500/10 text-emerald-400 border border-emerald-500/20">
                                    💵 كاش (7-)
                                  </span>
                                ) : (
                                  <span className="px-2 py-0.5 rounded-md text-[10px] font-black bg-blue-500/10 text-blue-400 border border-blue-500/20">
                                    💳 مدفوعات
                                  </span>
                                )}
                              </div>

                              {/* Card Content - Partners */}
                              <div className="space-y-1.5">
                                <span className="text-[10px] text-slate-400 font-bold block">
                                  {occ.totalRepsOnMachine === 1 ? 'الملكية والشركاء:' : 'المناديب المشاركون:'}
                                </span>
                                {occ.otherReps.length > 0 ? (
                                  <div className="flex items-center gap-1.5 flex-wrap">
                                    {occ.otherReps.map((r, rIdx) => (
                                      <span
                                        key={rIdx}
                                        onClick={() => setSearchQueryInstant(r.account !== 'غير متوفر' ? r.account : r.name)}
                                        className="px-2 py-0.5 rounded-lg bg-[#050811] hover:bg-slate-800 text-slate-300 border border-slate-800 hover:border-blue-500 text-[10px] hover:text-blue-400 cursor-pointer transition font-bold"
                                        title="اضغط للاستعلام عن هذا المندوب"
                                      >
                                        {r.name} {r.account !== 'غير متوفر' ? `(${r.account})` : ''} {r.type === 'cash' ? '💵 كاش' : '💳 دفع'}
                                      </span>
                                    ))}
                                  </div>
                                ) : occ.isRepeatedForSameRep ? (
                                  <span className="text-purple-300 text-[11px] font-bold">
                                    مسندة لنفس المندوب ({occ.totalForMachine}) مرات متكررة
                                  </span>
                                ) : (
                                  <span className="text-slate-500 text-[11px]">المندوب هو المالك الوحيد ولا يوجد شركاء</span>
                                )}
                              </div>

                              {/* Card Footer Actions */}
                              <div className="pt-2 border-t border-slate-800 flex items-center justify-between">
                                <span className="text-[10px] text-slate-500 font-mono font-bold">
                                  {occ.totalRepsOnMachine === 1 ? '1 مندوب' : `${occ.totalRepsOnMachine} مناديب`}
                                </span>
                                <button
                                  onClick={() => {
                                    setActiveTab('machine');
                                    setSearchQueryInstant(occ.machine);
                                  }}
                                  className="px-2.5 py-1 rounded-lg bg-indigo-600 hover:bg-indigo-500 text-white text-[10px] font-black transition cursor-pointer"
                                >
                                  فحص الماكينة &larr;
                                </button>
                              </div>
                            </div>
                          ))
                        ) : (
                          /* Summary Unique Cards */
                          filteredRepMachines.slice(0, visibleMachinesLimit).map((m, idx) => (
                            <div
                              key={`${m.machine}-${idx}`}
                              className={`border rounded-2xl p-4 flex flex-col justify-between gap-3 shadow-md transition-all ${
                                m.isRepeatedForSameRep
                                  ? 'bg-gradient-to-b from-[#18132e] to-[#0b1222] border-purple-500/50 shadow-purple-950/40 hover:border-purple-400'
                                  : 'bg-[#0b1222] border-slate-800 hover:border-blue-500/50'
                              }`}
                            >
                              {/* Card Header */}
                              <div className="flex items-center justify-between border-b border-slate-800 pb-2">
                                <div className="flex items-center gap-2">
                                  <span className="text-slate-500 font-mono text-[11px] font-bold">#{idx + 1}</span>
                                  <span className={`font-mono font-black text-sm tracking-wide ${m.type === 'cash' ? 'text-emerald-400' : 'text-white'}`}>
                                    {m.machine}
                                  </span>
                                </div>
                                <div className="flex flex-col items-end gap-1">
                                  <span
                                    className={`px-2 py-0.5 rounded-lg text-[10px] font-black border ${
                                      m.totalRepsOnMachine === 1
                                        ? 'bg-emerald-500/15 text-emerald-400 border-emerald-500/30'
                                        : 'bg-amber-500/15 text-amber-400 border-amber-500/30'
                                    }`}
                                  >
                                    {m.totalRepsOnMachine === 1 ? 'خاصة بالمندوب' : `مشتركة (${m.totalRepsOnMachine} مناديب)`}
                                  </span>
                                </div>
                              </div>

                              {/* Repetition and Type Badges */}
                              <div className="flex items-center gap-1.5 flex-wrap">
                                {m.isRepeatedForSameRep ? (
                                  <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-lg text-[10px] font-black bg-purple-500/20 text-purple-300 border border-purple-500/40 shadow-sm animate-pulse">
                                    <Repeat className="w-3 h-3" />
                                    <span>مكررة للمندوب ({m.sameRepOccurrences} مرات)</span>
                                  </span>
                                ) : (
                                  <span className="px-2 py-0.5 rounded-md text-[10px] font-bold bg-slate-800 text-slate-400 border border-slate-700">
                                    مرة واحدة
                                  </span>
                                )}

                                {m.type === 'both' ? (
                                  <span className="px-2 py-0.5 rounded-md text-[10px] font-black bg-amber-500/15 text-amber-300 border border-amber-500/30">
                                    ⭐ كاش + دفع
                                  </span>
                                ) : m.type === 'cash' ? (
                                  <span className="px-2 py-0.5 rounded-md text-[10px] font-black bg-emerald-500/10 text-emerald-400 border border-emerald-500/20">
                                    💵 كاش (7-)
                                  </span>
                                ) : (
                                  <span className="px-2 py-0.5 rounded-md text-[10px] font-black bg-blue-500/10 text-blue-400 border border-blue-500/20">
                                    💳 مدفوعات
                                  </span>
                                )}
                              </div>

                              {/* Repetition Description if repeated */}
                              {m.isRepeatedForSameRep && (
                                <div className="p-2 rounded-xl bg-purple-950/40 border border-purple-500/20 text-[10px] text-purple-200">
                                  <span className="font-bold block">تفاصيل التكرار:</span>
                                  <span>{m.repeatDescription}</span>
                                </div>
                              )}

                              {/* Card Content - Partners */}
                              <div className="space-y-1.5">
                                <span className="text-[10px] text-slate-400 font-bold block">
                                  {m.totalRepsOnMachine === 1 ? 'الملكية والشركاء:' : 'المناديب المشاركون:'}
                                </span>
                                {m.otherReps.length > 0 ? (
                                  <div className="flex items-center gap-1.5 flex-wrap">
                                    {m.otherReps.map((r, rIdx) => (
                                      <span
                                        key={rIdx}
                                        onClick={() => setSearchQueryInstant(r.account !== 'غير متوفر' ? r.account : r.name)}
                                        className="px-2 py-0.5 rounded-lg bg-[#050811] hover:bg-slate-800 text-slate-300 border border-slate-800 hover:border-blue-500 text-[10px] hover:text-blue-400 cursor-pointer transition font-bold"
                                        title="اضغط للاستعلام عن هذا المندوب"
                                      >
                                        {r.name} {r.account !== 'غير متوفر' ? `(${r.account})` : ''} {r.type === 'cash' ? '💵 كاش' : '💳 دفع'}
                                      </span>
                                    ))}
                                  </div>
                                ) : m.isRepeatedForSameRep ? (
                                  <span className="text-purple-300 text-[11px] font-bold">
                                    مسندة لنفس المندوب ({m.sameRepOccurrences}) مرات متكررة
                                  </span>
                                ) : (
                                  <span className="text-slate-500 text-[11px]">المندوب هو المالك الوحيد ولا يوجد شركاء</span>
                                )}
                              </div>

                              {/* Card Footer Actions */}
                              <div className="pt-2 border-t border-slate-800 flex items-center justify-between">
                                <span className="text-[10px] text-slate-500 font-mono font-bold">
                                  {m.totalRepsOnMachine === 1 ? '1 مندوب' : `${m.totalRepsOnMachine} مناديب`}
                                </span>
                                <button
                                  onClick={() => {
                                    setActiveTab('machine');
                                    setSearchQueryInstant(m.machine);
                                  }}
                                  className="px-2.5 py-1 rounded-lg bg-indigo-600 hover:bg-indigo-500 text-white text-[10px] font-black transition cursor-pointer"
                                >
                                  فحص الماكينة &larr;
                                </button>
                              </div>
                            </div>
                          ))
                        )}
                      </div>
                      {(repListDisplayMode === 'occurrences' ? filteredRepOccurrences.length : filteredRepMachines.length) > visibleMachinesLimit && (
                        <div className="p-3 text-center bg-[#0b1222] border border-slate-800 rounded-2xl mt-3">
                          <button
                            type="button"
                            onClick={() => setVisibleMachinesLimit((prev) => prev + 100)}
                            className="px-4 py-2 rounded-xl bg-blue-600/20 hover:bg-blue-600 text-blue-400 hover:text-white border border-blue-500/30 text-xs font-bold transition cursor-pointer"
                          >
                            عرض المزيد من السجلات (
                            {(repListDisplayMode === 'occurrences' ? filteredRepOccurrences.length : filteredRepMachines.length) - visibleMachinesLimit} متبقية من إجمالي{' '}
                            {repListDisplayMode === 'occurrences' ? filteredRepOccurrences.length : filteredRepMachines.length}
                            )
                          </button>
                        </div>
                      )}
                    </div>
                  )}
                </div>
              </div>
            ) : (
              <div className="py-12 text-center space-y-3 bg-[#090e18] rounded-2xl border border-slate-800 p-8">
                <AlertCircle className="w-10 h-10 text-amber-500 mx-auto" />
                <h4 className="font-bold text-white text-sm">
                  لم يتم العثور على مندوب يطابق: «{searchQuery}»
                </h4>
                <p className="text-xs text-slate-400 max-w-md mx-auto">
                  يرجى التأكد من كتابة رقم الحساب أو الاسم بشكل صحيح، أو اختر من القائمة المنسدلة بالأعلى.
                </p>
              </div>
            )
          )}

          {/* STATE C: MACHINE LOOKUP RESULTS */}
          {searchQuery.trim() && searchQuery === debouncedSearchQuery && activeTab === 'machine' && (
            matchedMachineData ? (
              <div className="space-y-4 animate-in fade-in duration-200">
                {/* Machine Profile Card */}
                <div className="p-4 rounded-2xl bg-[#0b1222] border border-slate-800 shadow-md">
                  <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
                    {/* Machine Info */}
                    <div className="flex items-center gap-3.5">
                      <div className="w-12 h-12 rounded-2xl bg-indigo-600/20 border border-indigo-500/30 text-indigo-400 flex items-center justify-center text-xl font-bold">
                        <Smartphone className="w-6 h-6" />
                      </div>
                      <div>
                        <div className="flex items-center gap-2 flex-wrap">
                          <h3 className="text-base font-black font-mono text-white tracking-wide">
                            ماكينة: {matchedMachineData.machine}
                          </h3>
                          <span
                            className={`text-xs font-bold px-2.5 py-0.5 rounded-lg border ${
                              matchedMachineData.status === 'single'
                                ? 'bg-emerald-500/10 text-emerald-400 border-emerald-500/30'
                                : matchedMachineData.status === 'multi'
                                ? 'bg-amber-500/10 text-amber-400 border-amber-500/30'
                                : 'bg-rose-500/10 text-rose-400 border-rose-500/30'
                            }`}
                          >
                            {matchedMachineData.status === 'single'
                              ? 'مندوب واحد (1/1)'
                              : matchedMachineData.status === 'multi'
                              ? `مشتركة (${matchedMachineData.repCount} مناديب)`
                              : 'شاغرة بدون مندوب'}
                          </span>
                        </div>
                        <p className="text-xs text-slate-400 mt-1">
                          تقرير تفصيلي بكافة أرقام حسابات وأسماء المناديب المربوطين بهذا الجهاز.
                        </p>
                      </div>
                    </div>

                    {/* KPI Pills */}
                    <div className="flex items-center gap-2 flex-wrap w-full md:w-auto justify-end">
                      <div className="px-4 py-2 rounded-xl bg-[#050811] border border-slate-800 text-center min-w-[100px]">
                        <span className="text-[10px] text-slate-400 font-bold block">عدد المناديب</span>
                        <span className="text-base font-mono font-black text-indigo-400">
                          {matchedMachineData.repCount}
                        </span>
                      </div>
                      <div className="px-4 py-2 rounded-xl bg-[#050811] border border-slate-800 text-center min-w-[100px]">
                        <span className="text-[10px] text-slate-400 font-bold block">حالة الماكينة</span>
                        <span className="text-xs font-bold text-white">
                          {matchedMachineData.repCount === 1 ? 'منفردة' : matchedMachineData.repCount > 1 ? 'مشتركة' : 'شاغرة'}
                        </span>
                      </div>
                    </div>
                  </div>
                </div>

                {/* Reps Cards for this Machine */}
                <div className="space-y-3">
                  <div className="p-3 bg-[#060a14] border border-slate-800 rounded-2xl flex items-center justify-between gap-3">
                    <div className="flex items-center gap-2">
                      <Users className="w-4 h-4 text-indigo-400" />
                      <span className="text-xs font-bold text-white">
                        المناديب المسندة لهذه الماكينة ({matchedMachineData.reps.length})
                      </span>
                    </div>

                    {/* Excel Export Button */}
                    <button
                      onClick={exportMachineStatement}
                      className="px-3 py-1.5 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold flex items-center gap-1.5 transition cursor-pointer shadow-md"
                    >
                      <FileSpreadsheet className="w-3.5 h-3.5" />
                      <span>تصدير بيان الماكينة Excel</span>
                    </button>
                  </div>

                  {/* Grid of Cards */}
                  <div className="overflow-y-auto max-h-[45vh] pr-1">
                    {matchedMachineData.reps.length === 0 ? (
                      <div className="p-8 text-center text-rose-400 bg-[#090e1a] border border-slate-800 rounded-2xl font-bold text-xs">
                        هذه الماكينة شاغرة ولم يتم إسناد أي مندوب لها في الشيت المرفوع.
                      </div>
                    ) : (
                      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
                        {matchedMachineData.reps.map((r, idx) => (
                          <div
                            key={`${r.account}-${idx}`}
                            className="bg-[#0b1222] border border-slate-800 rounded-2xl p-4 flex flex-col justify-between gap-3 shadow-md hover:border-indigo-500/50 transition-all"
                          >
                            {/* Card Header */}
                            <div className="flex items-center justify-between border-b border-slate-800 pb-2">
                              <div className="flex items-center gap-2">
                                <span className="text-slate-500 font-mono text-[11px] font-bold">#{idx + 1}</span>
                                <span className="font-mono font-bold text-blue-400 text-xs">
                                  {r.account !== 'غير متوفر' ? `حساب: ${r.account}` : 'بدون حساب'}
                                </span>
                              </div>
                              <span
                                className={`px-2 py-0.5 rounded-lg text-[10px] font-black border ${
                                  matchedMachineData.repCount === 1
                                    ? 'bg-emerald-500/15 text-emerald-400 border-emerald-500/30'
                                    : 'bg-amber-500/15 text-amber-400 border-amber-500/30'
                                }`}
                              >
                                {matchedMachineData.repCount === 1 ? 'منفرد (الوحيد)' : 'مشترك'}
                              </span>
                            </div>

                            {/* Card Content - Rep Profile */}
                            <div className="flex items-center gap-2.5">
                              <div className="w-8 h-8 rounded-xl bg-indigo-500/15 border border-indigo-500/30 text-indigo-400 flex items-center justify-center font-bold text-xs shrink-0">
                                <User className="w-4 h-4 text-indigo-400" />
                              </div>
                              <div className="min-w-0 flex-1">
                                <div className="flex items-center gap-1.5">
                                  <p className="font-black text-xs text-white truncate">{r.name}</p>
                                  <span className={`text-[9px] font-bold px-1.5 py-0.5 rounded border ${
                                    r.type === 'cash'
                                      ? 'bg-emerald-500/10 text-emerald-400 border-emerald-500/20'
                                      : 'bg-blue-500/10 text-blue-400 border-blue-500/20'
                                  }`}>
                                    {r.type === 'cash' ? '💵 كاش' : '💳 دفع'}
                                  </span>
                                </div>
                                <span className="text-[10px] text-slate-500 font-bold block mt-0.5">
                                  ترتيب التخصيص: مندوب {idx + 1} من {matchedMachineData.repCount}
                                </span>
                              </div>
                            </div>

                            {/* Card Footer Actions */}
                            <div className="pt-2 border-t border-slate-800 flex justify-end">
                              <button
                                onClick={() => {
                                  setActiveTab('rep');
                                  setSearchQueryInstant(r.account !== 'غير متوفر' ? r.account : r.name);
                                }}
                                className="px-2.5 py-1 rounded-lg bg-blue-600 hover:bg-blue-500 text-white text-[10px] font-black transition cursor-pointer"
                              >
                                كشف ماكيناته &larr;
                              </button>
                            </div>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                </div>
              </div>
            ) : (
              <div className="py-12 text-center space-y-3 bg-[#090e18] rounded-2xl border border-slate-800 p-8">
                <AlertCircle className="w-10 h-10 text-amber-500 mx-auto" />
                <h4 className="font-bold text-white text-sm">
                  لم يتم العثور على ماكينة تطابق: «{searchQuery}»
                </h4>
                <p className="text-xs text-slate-400 max-w-md mx-auto">
                  يرجى التأكد من كتابة رقم الماكينة بشكل صحيح (مثال: POS-100)، أو اختر من القائمة المنسدلة.
                </p>
              </div>
            )
          )}
        </div>

        {/* 5. FOOTER ACTIONS */}
        <div className="p-3 bg-[#0a0f1a] border-t border-slate-800/80 flex items-center justify-between gap-3">
          {/* Close button on the RIGHT in RTL */}
          <button
            onClick={onClose}
            className="bg-slate-800 hover:bg-slate-700 text-slate-300 font-bold text-xs px-4 py-2 rounded-xl border border-slate-700/60 transition cursor-pointer"
          >
            إغلاق
          </button>

          {/* Filter main table button on the LEFT in RTL */}
          {searchQuery && (
            <button
              onClick={() => {
                onFilterMainTable(searchQuery);
                onClose();
              }}
              className="bg-blue-600 hover:bg-blue-500 text-white font-bold text-xs px-4 py-2 rounded-xl shadow-lg shadow-blue-600/25 transition flex items-center gap-2 cursor-pointer"
            >
              <Filter className="w-3.5 h-3.5" />
              <span>تطبيق البحث على الجدول الرئيسي</span>
            </button>
          )}
        </div>
      </div>
    </div>
  );
};
