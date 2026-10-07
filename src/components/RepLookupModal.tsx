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

function sumFinancialColumns(rows: SheetRow[], keywords: string[]): number {
  if (!rows || rows.length === 0) return 0;
  
  // Find all columns in the first row that match any of the keywords
  const firstRow = rows[0];
  const matchingCols = Object.keys(firstRow).filter(key => {
    const normKey = key.toLowerCase();
    return keywords.some(kw => normKey.includes(kw.toLowerCase()));
  });

  if (matchingCols.length === 0) {
    // Fallback: search for any column that has purely numeric/decimal values except columns that look like IDs/Accounts
    const numericCols = Object.keys(firstRow).filter(key => {
      const normKey = key.toLowerCase();
      // Exclude IDs, machine numbers, accounts, telephone numbers, order index etc.
      if (/ماكينة|جهاز|pos|machine|sn|id|account|acc|code|تليفون|هاتف|رقم|موبايل|phone|mobile|date|تاريخ|index|serial|مسلسل|الترتيب/i.test(normKey)) {
        return false;
      }
      // Check if values are mostly numeric
      const val = firstRow[key];
      if (val === undefined || val === null) return false;
      const cleanVal = String(val).replace(/,/g, '').trim();
      const num = parseFloat(cleanVal);
      return !isNaN(num) && num > 0;
    });
    matchingCols.push(...numericCols);
  }

  // De-duplicate matching columns
  const finalCols = Array.from(new Set(matchingCols));

  let totalSum = 0;
  rows.forEach(row => {
    finalCols.forEach(col => {
      const val = row[col];
      if (val !== undefined && val !== null) {
        // Remove commas and parse
        const cleanVal = String(val).replace(/,/g, '').trim();
        const num = parseFloat(cleanVal);
        if (!isNaN(num)) {
          totalSum += num;
        }
      }
    });
  });

  return totalSum;
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
  const [searchQuery, setSearchQuery] = useState(initialRepQuery || '');

  // Reset or set search query when modal opens
  useEffect(() => {
    if (isOpen) {
      setSearchQuery(initialRepQuery || '');
      setActiveTab(initialMode || 'rep');
    }
  }, [isOpen, initialRepQuery, initialMode]);

  // Index Sheets directly (supports fast lookup even before full matching)
  const sheet2Index = useMemo(() => {
    const machineToRepsMap = new Map<string, Array<{ account: string; name: string; type: 'payment' | 'cash' }>>();
    const repToMachinesMap = new Map<string, { account: string; name: string; machines: Array<{ machine: string; type: 'payment' | 'cash' }> }>();

    if ((!sheet2 || sheet2.length === 0) && (!sheet3 || sheet3.length === 0)) {
      return {
        repList: [],
        machineList: [],
        machineToRepsMap,
        repToMachinesMap,
      };
    }

    // Process Sheet 2 (Payments)
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

      sheet2.forEach((row) => {
        const rawM = row[mCol];
        if (rawM === undefined || rawM === null || String(rawM).trim() === '') return;
        const rawMStr = String(rawM).trim();
        const normalizedKey = normalizeMachineId(rawMStr);
        const displayMachine = normalizedKey; // Payments machines are strictly clean without 7-

        const rawRep = row[repCol];
        const rawAcc = accCol && row[accCol] !== undefined ? String(row[accCol]).trim() : '';
        
        const repName = rawRep !== undefined && rawRep !== null && String(rawRep).trim() !== ''
          ? String(rawRep).trim()
          : (rawAcc ? `مندوب (${rawAcc})` : 'لا يوجد مندوب');

        const accVal = rawAcc || 'غير متوفر';

        // Index by machine
        if (!machineToRepsMap.has(normalizedKey)) {
          machineToRepsMap.set(normalizedKey, []);
        }
        machineToRepsMap.get(normalizedKey)!.push({ account: accVal, name: repName, type: 'payment' });

        // Index by rep (preserve every instance and unify across sheets)
        if (accVal !== 'غير متوفر' || (repName && repName !== 'لا يوجد مندوب')) {
          const normAccKey = accVal !== 'غير متوفر' ? normalizeText(accVal) : '';
          const normNameKey = repName ? normalizeText(repName) : '';
          const repKey = normAccKey || normNameKey;

          if (!repToMachinesMap.has(repKey)) {
            repToMachinesMap.set(repKey, { account: accVal, name: repName, machines: [] });
          }
          const repInfo = repToMachinesMap.get(repKey)!;
          if (accVal !== 'غير متوفر' && (repInfo.account === 'غير متوفر' || !repInfo.account)) {
            repInfo.account = accVal;
          }
          if (repName && !repName.startsWith('مندوب (') && !repName.startsWith('مسؤول كاش (')) {
            repInfo.name = repName;
          }
          repInfo.machines.push({ machine: displayMachine, type: 'payment' });

          // Also map secondary key if both account and name exist
          if (normAccKey && normNameKey && normNameKey !== normAccKey) {
            if (!repToMachinesMap.has(normNameKey)) {
              repToMachinesMap.set(normNameKey, repInfo);
            }
          }
        }
      });
    }

    // Process Sheet 3 (Cash)
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

      sheet3.forEach((row) => {
        const rawM = row[mCol3];
        if (rawM === undefined || rawM === null || String(rawM).trim() === '') return;
        const rawMStr = String(rawM).trim();
        const normalizedKey = normalizeMachineId(rawMStr);
        const displayMachine = rawMStr.startsWith('7-') || rawMStr.startsWith('٧-') ? rawMStr : `7-${normalizedKey}`;

        const rawRep = row[repCol3];
        const rawAcc = accCol3 && row[accCol3] !== undefined ? String(row[accCol3]).trim() : '';

        const repName = rawRep !== undefined && rawRep !== null && String(rawRep).trim() !== ''
          ? String(rawRep).trim()
          : (rawAcc ? `مسؤول كاش (${rawAcc})` : 'لا يوجد مسؤول كاش');

        const accVal = rawAcc || 'غير متوفر';

        // Index by machine
        if (!machineToRepsMap.has(normalizedKey)) {
          machineToRepsMap.set(normalizedKey, []);
        }
        machineToRepsMap.get(normalizedKey)!.push({ account: accVal, name: repName, type: 'cash' });

        // Index by rep (preserve every instance and unify with payments)
        if (accVal !== 'غير متوفر' || (repName && repName !== 'لا يوجد مسؤول كاش')) {
          const normAccKey = accVal !== 'غير متوفر' ? normalizeText(accVal) : '';
          const normNameKey = repName ? normalizeText(repName) : '';
          const repKey = normAccKey || normNameKey;

          if (!repToMachinesMap.has(repKey)) {
            repToMachinesMap.set(repKey, { account: accVal, name: repName, machines: [] });
          }
          const repInfo = repToMachinesMap.get(repKey)!;
          if (accVal !== 'غير متوفر' && (repInfo.account === 'غير متوفر' || !repInfo.account)) {
            repInfo.account = accVal;
          }
          if (repName && !repName.startsWith('مسؤول كاش (') && !repName.startsWith('مندوب (')) {
            repInfo.name = repName;
          }
          repInfo.machines.push({ machine: displayMachine, type: 'cash' });

          // Also map secondary key if both account and name exist
          if (normAccKey && normNameKey && normNameKey !== normAccKey) {
            if (!repToMachinesMap.has(normNameKey)) {
              repToMachinesMap.set(normNameKey, repInfo);
            }
          }
        }
      });
    }

    const uniqueRepValues = Array.from(new Set(Array.from(repToMachinesMap.values())));
    const repList = uniqueRepValues.map((r) => ({
      account: r.account,
      name: r.name,
      machineCount: r.machines.length,
      type: r.machines.some((m) => m.type === 'cash') ? ('cash' as const) : ('payment' as const),
    })).sort((a, b) => b.machineCount - a.machineCount);

    const machineList = Array.from(machineToRepsMap.entries()).map(([mKey, reps]) => {
      const norm = normalizeMachineId(mKey);
      const isPureCash = reps.length > 0 && reps.every(r => r.type === 'cash');
      const displayMachine = isPureCash 
        ? (norm.startsWith('7-') || norm.startsWith('٧-') ? norm : `7-${norm}`)
        : norm;
      return {
        machine: displayMachine,
        repCount: reps.length,
        reps,
        status: reps.length === 1 ? 'single' : reps.length > 1 ? 'multi' : 'none' as 'single' | 'multi' | 'none',
      };
    }).sort((a, b) => b.repCount - a.repCount);

    return { repList, machineList, machineToRepsMap, repToMachinesMap };
  }, [sheet2, sheet3, colM2, colM3, colAcc, colRep, colCashAcc, colCashRep]);

  // Unique Reps for dropdown & chips
  const uniqueReps = useMemo(() => {
    if (expandedRows && expandedRows.length > 0) {
      const map = new Map<string, { account: string; name: string; machineCount: number }>();
      expandedRows.forEach((row) => {
        if (row.status === 'none') return;
        const accKey = row.account.trim().toLowerCase();
        const nameVal = row.repName.trim();
        if (accKey === 'غير متوفر' && nameVal === 'لا يوجد مندوب') return;

        if (!map.has(accKey)) {
          map.set(accKey, { account: row.account, name: nameVal, machineCount: 0 });
        }
        map.get(accKey)!.machineCount += 1;
      });
      return Array.from(map.values()).sort((a, b) => b.machineCount - a.machineCount);
    }
    return sheet2Index.repList;
  }, [expandedRows, sheet2Index]);

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
    if (!searchQuery.trim()) {
      return uniqueReps.slice(0, 100); // Show top 100 by machine count
    }
    const q = searchQuery.toLowerCase().trim();
    return uniqueReps
      .filter((r) => r.name.toLowerCase().includes(q) || r.account.toLowerCase().includes(q))
      .slice(0, 100);
  }, [uniqueReps, searchQuery]);

  const filteredDropdownMachines = useMemo(() => {
    if (!searchQuery.trim()) {
      return uniqueMachines.slice(0, 100);
    }
    const q = searchQuery.toLowerCase().trim();
    return uniqueMachines
      .filter((m) => m.machine.toLowerCase().includes(q))
      .slice(0, 100);
  }, [uniqueMachines, searchQuery]);

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

  // Representative Matched Data
  const matchedRepData = useMemo(() => {
    if (!searchQuery.trim() || activeTab !== 'rep') {
      return null;
    }

    const q = searchQuery.trim().toLowerCase();
    const qNorm = q
      .replace(/[\u200B-\u200D\uFEFF\u00A0]/g, '')
      .replace(/[٠-٩]/g, (d) => String('٠١٢٣٤٥٦٧٨٩'.indexOf(d)));

    // Gather all matching reps from sheet2Index.repToMachinesMap
    const matchedEntries: Array<{
      account: string;
      name: string;
      machines: Array<{ machine: string; type: 'payment' | 'cash' }>;
    }> = [];

    const matchedRepKeys = new Set<string>();

    for (const [key, val] of sheet2Index.repToMachinesMap.entries()) {
      const keyNorm = key
        .replace(/[\u200B-\u200D\uFEFF\u00A0]/g, '')
        .replace(/[٠-٩]/g, (d) => String('٠١٢٣٤٥٦٧٨٩'.indexOf(d)));
      const accNorm = val.account
        .toLowerCase()
        .replace(/[\u200B-\u200D\uFEFF\u00A0]/g, '')
        .replace(/[٠-٩]/g, (d) => String('٠١٢٣٤٥٦٧٨٩'.indexOf(d)));
      const nameNorm = val.name
        .toLowerCase()
        .replace(/[\u200B-\u200D\uFEFF\u00A0]/g, '')
        .replace(/[٠-٩]/g, (d) => String('٠١٢٣٤٥٦٧٨٩'.indexOf(d)));

      if (keyNorm.includes(qNorm) || accNorm.includes(qNorm) || nameNorm.includes(qNorm) || qNorm.includes(accNorm)) {
        if (!matchedEntries.includes(val)) {
          matchedEntries.push(val);
        }
        if (val.account && val.account !== 'غير متوفر') matchedRepKeys.add(val.account.toLowerCase());
        if (val.name && val.name !== 'لا يوجد مندوب') matchedRepKeys.add(val.name.toLowerCase());
      }
    }

    // Also check expandedRows
    const matchedExpandedRows = (expandedRows || []).filter((r) => {
      if (r.status === 'none') return false;
      const accNorm = r.account
        .toLowerCase()
        .replace(/[\u200B-\u200D\uFEFF\u00A0]/g, '')
        .replace(/[٠-٩]/g, (d) => String('٠١٢٣٤٥٦٧٨٩'.indexOf(d)));
      const nameNorm = r.repName
        .toLowerCase()
        .replace(/[\u200B-\u200D\uFEFF\u00A0]/g, '')
        .replace(/[٠-٩]/g, (d) => String('٠١٢٣٤٥٦٧٨٩'.indexOf(d)));

      const isMatch =
        accNorm.includes(qNorm) ||
        nameNorm.includes(qNorm) ||
        matchedRepKeys.has(r.account.toLowerCase()) ||
        matchedRepKeys.has(r.repName.toLowerCase());

      if (isMatch) {
        if (r.account && r.account !== 'غير متوفر') matchedRepKeys.add(r.account.toLowerCase());
        if (r.repName && r.repName !== 'لا يوجد مندوب') matchedRepKeys.add(r.repName.toLowerCase());
      }
      return isMatch;
    });

    if (matchedEntries.length === 0 && matchedExpandedRows.length === 0) {
      return null;
    }

    // Determine primary rep name and account
    const primaryAccount =
      matchedEntries.find((e) => e.account && e.account !== 'غير متوفر')?.account ||
      matchedExpandedRows.find((r) => r.account && r.account !== 'غير متوفر')?.account ||
      searchQuery.trim();

    const primaryName =
      matchedEntries.find((e) => e.name && e.name !== 'لا يوجد مندوب' && e.name !== 'لا يوجد مسؤول كاش')?.name ||
      matchedExpandedRows.find((r) => r.repName && r.repName !== 'لا يوجد مندوب')?.repName ||
      `مندوب (${primaryAccount})`;

    // Build the complete list of machines for this rep (Both Cash and Payments!)
    const machinesList: Array<{
      machine: string;
      type: 'payment' | 'cash';
      totalRepsOnMachine: number;
      sameRepOccurrences: number;
      repAccount: string;
      repName: string;
      otherReps: Array<{ account: string; name: string; type?: 'payment' | 'cash' }>;
      status: 'single' | 'multi' | 'none';
    }> = [];

    const seenMachineKey = new Set<string>();

    // 1. First add from matchedExpandedRows (reconciled records with Sheet 1)
    matchedExpandedRows.forEach((row) => {
      const cleanId = normalizeMachineId(row.machine);
      const isCash = row.type === 'cash';
      const displayMachine = isCash
        ? (cleanId.startsWith('7-') || cleanId.startsWith('٧-') ? cleanId : `7-${cleanId}`)
        : cleanId;
      const dedupKey = `${displayMachine}__${row.type}__${row.id}`;

      if (!seenMachineKey.has(dedupKey)) {
        seenMachineKey.add(dedupKey);

        const normKey = normalizeMachineId(row.machine);
        const mSummary = (machinesResults || []).find((mr) => normalizeMachineId(mr.machine) === normKey);
        const allRepsOnMachine = mSummary?.reps || sheet2Index.machineToRepsMap.get(normKey) || [];

        const cleanCurrentAcc = normalizeText(row.account);
        const sameRepOccurrences = allRepsOnMachine.filter(
          (r) => (cleanCurrentAcc && cleanCurrentAcc !== 'غير متوفر' && normalizeText(r.account) === cleanCurrentAcc) ||
                 (r.name && r.name === row.repName)
        ).length;

        const otherReps = allRepsOnMachine.filter(
          (r) => (cleanCurrentAcc && cleanCurrentAcc !== 'غير متوفر' ? normalizeText(r.account) !== cleanCurrentAcc : r.name !== row.repName)
        );

        machinesList.push({
          machine: displayMachine,
          type: isCash ? 'cash' : 'payment',
          totalRepsOnMachine: Math.max(row.totalRepsForMachine, allRepsOnMachine.length, 1),
          sameRepOccurrences: Math.max(sameRepOccurrences, 1),
          repAccount: row.account,
          repName: row.repName,
          otherReps,
          status: row.status,
        });
      }
    });

    // 2. Also add any machines from matchedEntries (from Sheet 2 and Sheet 3) not already in machinesList
    matchedEntries.forEach((entry) => {
      entry.machines.forEach((mObj) => {
        const cleanId = normalizeMachineId(mObj.machine);
        const isCash = mObj.type === 'cash';
        const displayMachine = isCash
          ? (cleanId.startsWith('7-') || cleanId.startsWith('٧-') ? cleanId : `7-${cleanId}`)
          : cleanId;

        // Check if already included in machinesList
        const alreadyIncluded = machinesList.some(
          (m) => normalizeMachineId(m.machine) === cleanId && m.type === mObj.type
        );

        if (!alreadyIncluded) {
          const normKey = normalizeMachineId(mObj.machine);
          const allRepsOnMachine = sheet2Index.machineToRepsMap.get(normKey) || [
            { account: entry.account, name: entry.name, type: mObj.type }
          ];

          const cleanCurrentAcc = normalizeText(entry.account);
          const sameRepOccurrences = allRepsOnMachine.filter(
            (r) => (cleanCurrentAcc && cleanCurrentAcc !== 'غير متوفر' && normalizeText(r.account) === cleanCurrentAcc) ||
                   (r.name && r.name === entry.name)
          ).length;

          const otherReps = allRepsOnMachine.filter(
            (r) => (cleanCurrentAcc && cleanCurrentAcc !== 'غير متوفر' ? normalizeText(r.account) !== cleanCurrentAcc : r.name !== entry.name)
          );

          machinesList.push({
            machine: displayMachine,
            type: isCash ? 'cash' : 'payment',
            totalRepsOnMachine: Math.max(allRepsOnMachine.length, 1),
            sameRepOccurrences: Math.max(sameRepOccurrences, 1),
            repAccount: entry.account,
            repName: entry.name,
            otherReps,
            status: allRepsOnMachine.length > 1 ? 'multi' : 'single',
          });
        }
      });
    });

    const cashCount = machinesList.filter((m) => m.type === 'cash').length;
    const paymentsCount = machinesList.filter((m) => m.type === 'payment').length;
    const soleCount = machinesList.filter((m) => m.totalRepsOnMachine === 1).length;
    const sharedCount = machinesList.filter((m) => m.totalRepsOnMachine > 1).length;

    // Filter sheet2 and sheet3 rows to calculate exact financial sums
    const matchedSheet2Rows = (sheet2 || []).filter(row => {
      const sampleRow = row;
      const keys = Object.keys(sampleRow);
      
      const accCol = colAcc && sampleRow[colAcc] !== undefined
        ? colAcc
        : keys.find((k) => /^(doner1|donor1|doner_1|donor_1|doner 1|donor 1|doner|donor|رقم حساب المدفوعات|حساب المدفوعات)$/i.test(k))
          || keys.find((k) => /(doner.*1|donor.*1|^doner$|^donor$|حساب.*مدفوعات|رقم.*المدفوعات)/i.test(k))
          || keys.find((k) => /حساب|كود|رقم.*مندوب|account|acc|code|doner|donor/i.test(k)) || '';

      const repCol = colRep && sampleRow[colRep] !== undefined
        ? colRep
        : keys.find((k) => /اسم.*مندوب|اسم.*العميل|اسم.*المستخدم|اسم|rep.*name|agent.*name|name/i.test(k))
          || accCol || keys[1] || keys[0];

      const accVal = accCol && row[accCol] !== undefined ? String(row[accCol]).trim() : '';
      const repName = row[repCol] !== undefined && row[repCol] !== null ? String(row[repCol]).trim() : '';

      const accNorm = accVal.toLowerCase().replace(/[\u200B-\u200D\uFEFF\u00A0]/g, '').replace(/[٠-٩]/g, (d) => String('٠١٢٣٤٥٦٧٨٩'.indexOf(d)));
      const nameNorm = repName.toLowerCase().replace(/[\u200B-\u200D\uFEFF\u00A0]/g, '').replace(/[٠-٩]/g, (d) => String('٠١٢٣٤٥٦٧٨٩'.indexOf(d)));

      return accNorm.includes(qNorm) || nameNorm.includes(qNorm) || qNorm.includes(accNorm);
    });

    const matchedSheet3Rows = (sheet3 || []).filter(row => {
      const sampleRow = row;
      const keys = Object.keys(sampleRow);

      const accCol3 = colCashAcc && sampleRow[colCashAcc] !== undefined
        ? colCashAcc
        : keys.find((k) => /^(doner2|donor2|doner_2|donor_2|doner 2|donor 2|رقم الكاش المدفوعات|رقم الكاش|حساب الكاش)$/i.test(k))
          || keys.find((k) => /(doner.*2|donor.*2|كاش.*مدفوعات|رقم.*الكاش|حساب.*الكاش)/i.test(k))
          || keys.find((k) => /حساب|كود|رقم.*مندوب|account|acc|code|doner|donor/i.test(k)) || '';

      const repCol3 = colCashRep && sampleRow[colCashRep] !== undefined
        ? colCashRep
        : keys.find((k) => /اسم.*مندوب|اسم.*العميل|اسم.*المستخدم|اسم|rep.*name|agent.*name|name/i.test(k))
          || accCol3 || keys[1] || keys[0];

      const accVal = accCol3 && row[accCol3] !== undefined ? String(row[accCol3]).trim() : '';
      const repName = row[repCol3] !== undefined && row[repCol3] !== null ? String(row[repCol3]).trim() : '';

      const accNorm = accVal.toLowerCase().replace(/[\u200B-\u200D\uFEFF\u00A0]/g, '').replace(/[٠-٩]/g, (d) => String('٠١٢٣٤٥٦٧٨٩'.indexOf(d)));
      const nameNorm = repName.toLowerCase().replace(/[\u200B-\u200D\uFEFF\u00A0]/g, '').replace(/[٠-٩]/g, (d) => String('٠١٢٣٤٥٦٧٨٩'.indexOf(d)));

      return accNorm.includes(qNorm) || nameNorm.includes(qNorm) || qNorm.includes(accNorm);
    });

    const totalPaymentsAmount = sumFinancialColumns(matchedSheet2Rows, ['مبلغ', 'قيمة', 'القيمة', 'رصيد', 'صافي', 'الصافي', 'المدفوع', 'المدفوعات', 'القيمه', 'حركة', 'الحركات', 'amount', 'value', 'total', 'net', 'pay', 'sum', 'balance']);
    const totalCashAmount = sumFinancialColumns(matchedSheet3Rows, ['مبلغ', 'قيمة', 'القيمة', 'رصيد', 'صافي', 'الصافي', 'الكاش', 'كاش', 'القيمه', 'حركة', 'الحركات', 'amount', 'value', 'total', 'net', 'cash', 'sum', 'balance']);

    return {
      repInfo: {
        account: primaryAccount,
        name: primaryName,
        totalMachines: machinesList.length,
        cashCount,
        paymentsCount,
        soleCount,
        sharedCount,
        totalPaymentsAmount,
        totalCashAmount,
        totalOverallAmount: totalPaymentsAmount + totalCashAmount,
      },
      machines: machinesList,
    };
  }, [searchQuery, activeTab, expandedRows, machinesResults, sheet2Index, sheet2, sheet3, colAcc, colRep, colCashAcc, colCashRep]);

  // Machine Matched Data (الاستعلام عن ماكينة معينة)
  const matchedMachineData = useMemo(() => {
    if (!searchQuery.trim() || activeTab !== 'machine') {
      return null;
    }

    const q = searchQuery.trim().toLowerCase();

    // CASE 1: Using machinesResults
    if (machinesResults && machinesResults.length > 0) {
      const qNorm = normalizeMachineId(q);
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

    // CASE 2: Using sheet2 index / sheet1
    if (sheet2Index.machineToRepsMap.size > 0 || sheet1.length > 0) {
      const qNorm = normalizeMachineId(q);
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
  }, [searchQuery, activeTab, machinesResults, sheet2Index, sheet1]);

  if (!isOpen) return null;

  // Total records count read
  const totalSheetRecords = sheet2.length || sheet1.length || expandedRows.length;
  const isSheetLoaded = totalSheetRecords > 0;

  // Export representative statement to Excel
  const exportRepStatement = () => {
    if (!matchedRepData || !matchedRepData.machines.length) return;

    const rep = matchedRepData.repInfo;
    const data: (string | number)[][] = [
      ['كشف الماكينات المربوطة بالمندوب'],
      ['اسم المندوب:', rep.name],
      ['رقم الحساب:', rep.account],
      ['إجمالي الماكينات المربوطة:', rep.totalMachines],
      ['ماكينات منفردة (خاصة بالمندوب فقط):', rep.soleCount],
      ['ماكينات مشتركة مع مناديب آخرين:', rep.sharedCount],
      [],
      ['م', 'رقم الماكينة', 'حالة الملكية والتعدد', 'الترتيب والتعدد', 'المناديب المشاركون على نفس الماكينة'],
    ];

    matchedRepData.machines.forEach((m, idx) => {
      const coRepsStr = m.otherReps.length > 0
        ? m.otherReps.map((r) => `${r.name} (${r.account})`).join(' ، ')
        : 'لا يوجد مناديب آخرون (المندوب الوحيد لهذه الماكينة)';

      const stStr = m.totalRepsOnMachine === 1 ? 'خاصة فقط' : `مشتركة (${m.totalRepsOnMachine} مناديب)`;

      data.push([idx + 1, m.machine, stStr, `${m.totalRepsOnMachine} مناديب`, coRepsStr]);
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [{ wch: 6 }, { wch: 22 }, { wch: 22 }, { wch: 18 }, { wch: 50 }];
    XLSX.utils.book_append_sheet(wb, ws, `ماكينات_${rep.account}`);
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
                setSearchQuery('');
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
                setSearchQuery('');
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
                  onClick={() => setSearchQuery('')}
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
                  onChange={(e) => setSearchQuery(e.target.value)}
                  className="w-full bg-[#050811] border border-slate-700/80 rounded-xl px-3 py-2.5 text-xs text-blue-400 font-bold focus:outline-none focus:border-blue-500 transition cursor-pointer shadow-inner truncate"
                >
                  <option value="">-- اختر مندوب من القائمة ({uniqueReps.length}) --</option>
                  {filteredDropdownReps.map((r, rIdx) => (
                    <option key={`${r.account}-${rIdx}`} value={r.account !== 'غير متوفر' ? r.account : r.name}>
                      {r.name} {r.account !== 'غير متوفر' ? `(حساب: ${r.account})` : ''} • [{r.machineCount} ماكينة]
                    </option>
                  ))}
                </select>
              ) : (
                <select
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
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
                      onClick={() => setSearchQuery(r.account !== 'غير متوفر' ? r.account : r.name)}
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
                      onClick={() => setSearchQuery(m.machine)}
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

          {/* STATE B: REP LOOKUP RESULTS */}
          {searchQuery.trim() && activeTab === 'rep' && (
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
                        </div>
                        <p className="text-xs text-slate-400 mt-1">
                          بيان تفصيلي بجميع ماكينات هذا المندوب وحالة الاشتراك والتعدد لكل جهاز.
                        </p>
                      </div>
                    </div>

                    {/* KPI Pills */}
                    <div className="flex items-center gap-2 flex-wrap w-full md:w-auto justify-end">
                      <div className="px-3.5 py-2 rounded-xl bg-[#050811] border border-slate-800 text-center min-w-[80px]">
                        <span className="text-[10px] text-slate-400 font-bold block">إجمالي الماكينات</span>
                        <span className="text-base font-mono font-black text-blue-400">
                          {matchedRepData.repInfo.totalMachines}
                        </span>
                      </div>

                      {/* Cash Machines Count Box */}
                      <div className="px-3.5 py-2 rounded-xl bg-[#050811] border border-emerald-500/40 text-center min-w-[95px] shadow-sm">
                        <span className="text-[10px] text-emerald-400 font-bold flex items-center justify-center gap-1">
                          <span>💵 ماكينات الكاش</span>
                        </span>
                        <span className="text-base font-mono font-black text-emerald-400">
                          {matchedRepData.repInfo.cashCount}
                        </span>
                      </div>

                      {/* Payments Machines Count Box */}
                      <div className="px-3.5 py-2 rounded-xl bg-[#050811] border border-blue-500/40 text-center min-w-[95px] shadow-sm">
                        <span className="text-[10px] text-blue-400 font-bold flex items-center justify-center gap-1">
                          <span>💳 ماكينات المدفوعات</span>
                        </span>
                        <span className="text-base font-mono font-black text-blue-400">
                          {matchedRepData.repInfo.paymentsCount}
                        </span>
                      </div>

                      <div className="px-3.5 py-2 rounded-xl bg-[#050811] border border-emerald-900/40 text-center min-w-[80px]">
                        <span className="text-[10px] text-emerald-400 font-bold block">منفردة (خاصة)</span>
                        <span className="text-base font-mono font-black text-emerald-400">
                          {matchedRepData.repInfo.soleCount}
                        </span>
                      </div>
                      <div className="px-3.5 py-2 rounded-xl bg-[#050811] border border-amber-900/40 text-center min-w-[80px]">
                        <span className="text-[10px] text-amber-400 font-bold block">مشتركة (متعددة)</span>
                        <span className="text-base font-mono font-black text-amber-400">
                          {matchedRepData.repInfo.sharedCount}
                        </span>
                      </div>
                    </div>
                  </div>

                  {/* Financial Statistics Row */}
                  {(matchedRepData.repInfo.totalPaymentsAmount > 0 || matchedRepData.repInfo.totalCashAmount > 0) && (
                    <div className="mt-4 pt-3.5 border-t border-slate-800/80 grid grid-cols-1 sm:grid-cols-3 gap-3 text-right">
                      {/* Cash Sum Card */}
                      <div className="p-3 rounded-xl bg-emerald-500/5 border border-emerald-500/20 flex items-center justify-between shadow-sm">
                        <div className="flex items-center gap-2">
                          <span className="text-lg">💵</span>
                          <div className="text-right">
                            <span className="text-[10px] text-emerald-400 font-bold block">إجمالي مبالغ الكاش</span>
                            <span className="text-[10px] text-slate-400">من واقع شيت الكاش</span>
                          </div>
                        </div>
                        <span className="text-sm sm:text-base font-mono font-black text-emerald-400">
                          {matchedRepData.repInfo.totalCashAmount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                        </span>
                      </div>

                      {/* Payments Sum Card */}
                      <div className="p-3 rounded-xl bg-blue-500/5 border border-blue-500/20 flex items-center justify-between shadow-sm">
                        <div className="flex items-center gap-2">
                          <span className="text-lg">💳</span>
                          <div className="text-right">
                            <span className="text-[10px] text-blue-400 font-bold block">إجمالي مبالغ المدفوعات</span>
                            <span className="text-[10px] text-slate-400">من واقع شيت المدفوعات</span>
                          </div>
                        </div>
                        <span className="text-sm sm:text-base font-mono font-black text-blue-400">
                          {matchedRepData.repInfo.totalPaymentsAmount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                        </span>
                      </div>

                      {/* Combined Sum Card */}
                      <div className="p-3 rounded-xl bg-indigo-500/10 border border-indigo-500/20 flex items-center justify-between shadow-sm">
                        <div className="flex items-center gap-2">
                          <span className="text-lg">💰</span>
                          <div className="text-right">
                            <span className="text-[10px] text-indigo-400 font-bold block">إجمالي المبالغ بالكامل</span>
                            <span className="text-[10px] text-slate-400">كاش + مدفوعات مندوب</span>
                          </div>
                        </div>
                        <span className="text-sm sm:text-base font-mono font-black text-white">
                          {matchedRepData.repInfo.totalOverallAmount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                        </span>
                      </div>
                    </div>
                  )}
                </div>

                {/* Machines Section for this Rep */}
                <div className="space-y-3">
                  <div className="p-3 bg-[#060a14] border border-slate-800 rounded-2xl flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                    <div className="flex items-center gap-2">
                      <Layers className="w-4 h-4 text-blue-400" />
                      <span className="text-xs font-bold text-white">
                        قائمة الماكينات المربوطة ({matchedRepData.machines.length})
                      </span>
                    </div>

                    <div className="flex items-center gap-2.5 flex-wrap">
                      {/* View Switcher Capsule (Table vs Cards) */}
                      <div className="inline-flex rounded-xl p-1 bg-[#050811] border border-slate-800 text-xs font-bold shrink-0">
                        <button
                          onClick={() => setLookupViewMode('cards')}
                          className={`px-3 py-1 rounded-lg flex items-center gap-1.5 transition-all cursor-pointer ${
                            lookupViewMode === 'cards'
                              ? 'bg-blue-600 text-white shadow-sm'
                              : 'text-slate-400 hover:text-white'
                          }`}
                          title="عرض البطاقات والبوكسات"
                        >
                          <Layers className="w-3.5 h-3.5" />
                          <span>عرض البطاقات</span>
                        </button>
                        <button
                          onClick={() => setLookupViewMode('table')}
                          className={`px-3 py-1 rounded-lg flex items-center gap-1.5 transition-all cursor-pointer ${
                            lookupViewMode === 'table'
                              ? 'bg-blue-600 text-white shadow-sm'
                              : 'text-slate-400 hover:text-white'
                          }`}
                          title="عرض جدول البيانات التفصيلي"
                        >
                          <TableProperties className="w-3.5 h-3.5" />
                          <span>عرض الجدول</span>
                        </button>
                      </div>

                      {/* Excel Export Button */}
                      <button
                        onClick={exportRepStatement}
                        className="px-3 py-1.5 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold flex items-center gap-1.5 transition cursor-pointer shadow-md"
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
                            <th className="p-3 w-40 font-mono">رقم الماكينة</th>
                            <th className="p-3 w-32 text-center">نوع الحساب</th>
                            <th className="p-3 w-40 text-center">حالة الملكية والتعدد</th>
                            <th className="p-3">المناديب المشاركون على الماكينة</th>
                            <th className="p-3 w-28 text-center">إجراء</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-800/80 font-mono text-slate-200">
                          {matchedRepData.machines.map((m, idx) => (
                            <tr key={`${m.machine}-${idx}`} className="hover:bg-slate-800/50 transition-colors">
                              <td className="p-3 text-center text-slate-500 font-bold">{idx + 1}</td>
                              <td className="p-3 font-bold text-white text-sm tracking-wide">
                                {m.machine}
                              </td>
                              <td className="p-3 text-center font-bold">
                                {m.type === 'cash' ? (
                                  <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[10px] font-black bg-emerald-500/10 text-emerald-400 border border-emerald-500/20">
                                    💵 كاش
                                  </span>
                                ) : (
                                  <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[10px] font-black bg-blue-500/10 text-blue-400 border border-blue-500/20">
                                    💳 مدفوعات
                                  </span>
                                )}
                              </td>
                              <td className="p-3 text-center">
                                <div className="flex flex-col items-center gap-1">
                                  <span
                                    className={`inline-block px-2 py-0.5 rounded-lg text-[10px] font-black border ${
                                      m.totalRepsOnMachine === 1
                                        ? 'bg-emerald-500/15 text-emerald-400 border-emerald-500/30'
                                        : 'bg-amber-500/15 text-amber-400 border-amber-500/30'
                                    }`}
                                  >
                                    {m.totalRepsOnMachine === 1 ? 'خاصة بالمندوب' : `مشتركة (${m.totalRepsOnMachine} مناديب)`}
                                  </span>
                                  {m.sameRepOccurrences > 1 && (
                                    <span className="inline-block px-1.5 py-0.5 rounded-md text-[9px] font-black bg-purple-500/15 text-purple-400 border border-purple-500/30">
                                      🔁 مكرر للمندوب ({m.sameRepOccurrences} مرات)
                                    </span>
                                  )}
                                </div>
                              </td>
                              <td className="p-3 font-sans">
                                {m.otherReps.length > 0 ? (
                                  <div className="flex items-center gap-1.5 flex-wrap">
                                    {m.otherReps.map((r, rIdx) => (
                                      <span
                                        key={rIdx}
                                        onClick={() => setSearchQuery(r.account !== 'غير متوفر' ? r.account : r.name)}
                                        className="px-2 py-0.5 rounded-lg bg-[#050811] hover:bg-slate-800 text-slate-300 border border-slate-800 hover:border-blue-500 text-[10px] hover:text-blue-400 cursor-pointer transition font-bold"
                                        title="اضغط للاستعلام عن هذا المندوب"
                                      >
                                        {r.name} {r.account !== 'غير متوفر' ? `(${r.account})` : ''} {r.type === 'cash' ? '💵' : '💳'}
                                      </span>
                                    ))}
                                  </div>
                                ) : m.sameRepOccurrences > 1 ? (
                                  <span className="text-purple-400 text-[11px] font-bold">
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
                                    setSearchQuery(m.machine);
                                  }}
                                  className="px-2.5 py-1 rounded-lg bg-indigo-600 hover:bg-indigo-500 text-white text-[10px] font-black transition cursor-pointer"
                                >
                                  فحص &larr;
                                </button>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  ) : (
                    /* Grid of Cards */
                    <div className="overflow-y-auto max-h-[45vh] pr-1">
                      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
                        {matchedRepData.machines.map((m, idx) => (
                          <div
                            key={`${m.machine}-${idx}`}
                            className="bg-[#0b1222] border border-slate-800 rounded-2xl p-4 flex flex-col justify-between gap-3 shadow-md hover:border-blue-500/50 transition-all"
                          >
                            {/* Card Header */}
                            <div className="flex items-center justify-between border-b border-slate-800 pb-2">
                              <div className="flex items-center gap-2">
                                <span className="text-slate-500 font-mono text-[11px] font-bold">#{idx + 1}</span>
                                <span className="font-mono font-black text-white text-sm tracking-wide">{m.machine}</span>
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
                                {m.sameRepOccurrences > 1 && (
                                  <span className="px-1.5 py-0.5 rounded-md text-[9px] font-black bg-purple-500/15 text-purple-400 border border-purple-500/30">
                                    🔁 مكرر ({m.sameRepOccurrences} مرات)
                                  </span>
                                )}
                              </div>
                            </div>

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
                                      onClick={() => setSearchQuery(r.account !== 'غير متوفر' ? r.account : r.name)}
                                      className="px-2 py-0.5 rounded-lg bg-[#050811] hover:bg-slate-800 text-slate-300 border border-slate-800 hover:border-blue-500 text-[10px] hover:text-blue-400 cursor-pointer transition font-bold"
                                      title="اضغط للاستعلام عن هذا المندوب"
                                    >
                                      {r.name} {r.account !== 'غير متوفر' ? `(${r.account})` : ''} {r.type === 'cash' ? '💵 كاش' : '💳 دفع'}
                                    </span>
                                  ))}
                                </div>
                              ) : m.sameRepOccurrences > 1 ? (
                                <span className="text-purple-400 text-[11px] font-bold">
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
                                  setSearchQuery(m.machine);
                                }}
                                className="px-2.5 py-1 rounded-lg bg-indigo-600 hover:bg-indigo-500 text-white text-[10px] font-black transition cursor-pointer"
                              >
                                فحص الماكينة &larr;
                              </button>
                            </div>
                          </div>
                        ))}
                      </div>
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
          {searchQuery.trim() && activeTab === 'machine' && (
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
                                  setSearchQuery(r.account !== 'غير متوفر' ? r.account : r.name);
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
