import * as XLSX from 'xlsx';
import { ExpandedRow, SheetRow } from '../types';

export function filterRowsForTarget(
  rows: SheetRow[],
  targetType?: 'machines' | 'payments' | 'cash'
): SheetRow[] {
  if (!rows || rows.length === 0 || !targetType) return rows;

  const sampleRow = rows[0];
  const allCols = Object.keys(sampleRow);

  // Identify acceptor2 / cash machine column
  const acceptor2Col = allCols.find(
    (c) =>
      /^(acceptor2|acceptor_2|acceptor 2|رقم ماكينه الكاش|ماكينه الكاش|ماكينة الكاش)$/i.test(c) ||
      /(acceptor.*2|ماكين.*كاش|ماكينة.*الكاش)/i.test(c)
  );

  // Identify doner2 / cash account column
  const doner2Col = allCols.find(
    (c) =>
      /^(doner2|donor2|doner_2|donor_2|doner 2|donor 2|رقم الكاش المدفوعات|رقم الكاش|حساب الكاش)$/i.test(c) ||
      /(doner.*2|donor.*2|كاش.*مدفوعات|رقم.*الكاش|حساب.*الكاش)/i.test(c)
  );

  // Identify acceptor / acceptor1 (payments machine)
  const acceptor1Col = allCols.find(
    (c) =>
      /^(acceptor1|acceptor_1|acceptor 1|acceptor|ماكينة المدفوعات|ماكينه المدفوعات)$/i.test(c) ||
      /(acceptor.*1|^acceptor$|ماكين.*مدفوعات|ماكينة.*المدفوعات)/i.test(c)
  );

  // Identify doner / doner1 (payments account)
  const doner1Col = allCols.find(
    (c) =>
      /^(doner1|donor1|doner_1|donor_1|doner 1|donor 1|doner|donor|رقم حساب المدفوعات|حساب المدفوعات)$/i.test(c) ||
      /(doner.*1|donor.*1|^doner$|^donor$|حساب.*مدفوعات|رقم.*المدفوعات)/i.test(c)
  );

  // TARGET: CASH
  if (targetType === 'cash') {
    // If the file has explicit acceptor2 or doner2 columns:
    // User requested: "بمعنى لما ارفع الشيت ف الكاش يظهر دول فقط: doner2 رقم الكاش المدفوعات، acceptor2 رقم ماكينه الكاش"
    if (acceptor2Col || doner2Col) {
      const result: SheetRow[] = [];
      for (const r of rows) {
        const mVal = acceptor2Col ? String(r[acceptor2Col] ?? '').trim() : '';
        const dVal = doner2Col ? String(r[doner2Col] ?? '').trim() : '';
        if (mVal !== '' || dVal !== '') {
          const rowObj: SheetRow = {};
          if (acceptor2Col) rowObj[acceptor2Col] = mVal;
          if (doner2Col) rowObj[doner2Col] = dVal;
          rowObj._tabType = 'cash';
          result.push(rowObj);
        }
      }
      if (result.length > 0) return result;
    }

    // If rows have _tabType, keep only cash rows
    const cashOnly = rows.filter(
      (r) => r._tabType === 'cash' || (r._tabType === undefined && r._tabType !== 'payments')
    );
    if (cashOnly.length > 0 && cashOnly.length < rows.length) {
      return cashOnly;
    }
  }

  // TARGET: PAYMENTS
  if (targetType === 'payments') {
    // If the file has payments columns and was a multi-column shared file
    if ((acceptor1Col || doner1Col) && (acceptor2Col || doner2Col)) {
      const result: SheetRow[] = [];
      for (const r of rows) {
        const mVal = acceptor1Col ? String(r[acceptor1Col] ?? '').trim() : '';
        const dVal = doner1Col ? String(r[doner1Col] ?? '').trim() : '';
        if (mVal !== '' || dVal !== '') {
          const rowObj: SheetRow = {};
          if (acceptor1Col) rowObj[acceptor1Col] = mVal;
          if (doner1Col) rowObj[doner1Col] = dVal;
          rowObj._tabType = 'payments';
          result.push(rowObj);
        }
      }
      if (result.length > 0) return result;
    }

    // If rows have _tabType, keep only payments rows
    const paymentsOnly = rows.filter(
      (r) => r._tabType === 'payments' || (r._tabType === undefined && r._tabType !== 'cash')
    );
    if (paymentsOnly.length > 0 && paymentsOnly.length < rows.length) {
      return paymentsOnly;
    }
  }

  return rows;
}

export async function parseExcelFile(
  file: File,
  targetType?: 'machines' | 'payments' | 'cash'
): Promise<SheetRow[]> {
  const isCsv = file.name.toLowerCase().endsWith('.csv');

  // Fast-path for CSV files using direct text processing
  if (isCsv) {
    try {
      const text = await file.text();
      const lines = text.split(/\r\n|\n|\r/);
      if (lines.length > 0) {
        let headerRowIdx = -1;
        let headers: string[] = [];

        for (let i = 0; i < Math.min(lines.length, 12); i++) {
          const line = lines[i].trim();
          if (!line) continue;
          const cols = line.split(',').map((c) => c.replace(/^["']|["']$/g, '').trim());
          if (cols.some((c) => /ماكينة|جهاز|pos|machine|sn|id|كود|حساب|مندوب|اسم|فرع|code|acc|name|acceptor|doner|donor/i.test(c))) {
            headerRowIdx = i;
            headers = cols;
            break;
          }
        }

        if (headerRowIdx === -1 && lines.length > 0) {
          headerRowIdx = 0;
          headers = lines[0].split(',').map((c, colIdx) => c.replace(/^["']|["']$/g, '').trim() || `عمود_${colIdx + 1}`);
        }

        const rows: SheetRow[] = [];
        for (let i = headerRowIdx + 1; i < lines.length; i++) {
          const line = lines[i].trim();
          if (!line) continue;
          const cols = line.split(',').map((c) => c.replace(/^["']|["']$/g, '').trim());
          const rowObj: SheetRow = {};
          let hasData = false;

          for (let c = 0; c < headers.length; c++) {
            const val = cols[c] !== undefined ? cols[c] : '';
            const h = headers[c];
            rowObj[h] = val;
            if (val !== '') hasData = true;
          }

          if (hasData) rows.push(rowObj);
        }

        if (rows.length > 0) {
          return filterRowsForTarget(rows, targetType);
        }
      }
    } catch (e) {
      // Fallback to XLSX parser if text parsing fails
    }
  }

  const buffer = await file.arrayBuffer();
  let workbook: XLSX.WorkBook | null = null;

  try {
    workbook = XLSX.read(buffer, {
      type: 'array',
      codepage: 65001,
      cellDates: false,
      cellStyles: false,
      cellFormula: false,
      sheetStubs: false,
      dense: true,
    });
  } catch (err1) {
    try {
      const text = new TextDecoder('utf-8').decode(buffer);
      workbook = XLSX.read(text, {
        type: 'string',
        cellDates: false,
        cellStyles: false,
        cellFormula: false,
        sheetStubs: false,
        dense: true,
      });
    } catch (err2) {
      try {
        const text1256 = new TextDecoder('windows-1256').decode(buffer);
        workbook = XLSX.read(text1256, {
          type: 'string',
          cellDates: false,
          cellStyles: false,
          cellFormula: false,
          sheetStubs: false,
          dense: true,
        });
      } catch (err3) {
        throw new Error('تعذر قراءة صيغة الملف، يرجى التأكد من صلاحية ملف الإكسل أو CSV.');
      }
    }
  }

  if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
    return [];
  }

  const extractRowsFromWorksheet = (worksheet: XLSX.WorkSheet): SheetRow[] => {
    if (!worksheet || !worksheet['!ref']) return [];
    const aoa = XLSX.utils.sheet_to_json<any[]>(worksheet, { header: 1, defval: '', raw: true });
    if (!aoa || aoa.length === 0) return [];

    let headerRowIdx = -1;
    for (let r = 0; r < Math.min(aoa.length, 12); r++) {
      const row = aoa[r];
      if (!Array.isArray(row)) continue;
      const nonEmptyCells = row.map((c) => String(c ?? '').trim()).filter((c) => c !== '');
      if (nonEmptyCells.length >= 1) {
        const hasKnownKeywords = nonEmptyCells.some((c) =>
          /ماكينة|جهاز|pos|machine|sn|id|كود|حساب|مندوب|اسم|فرع|تاريخ|مسلسل|code|acc|name|terminal|merchant|acceptor|doner|donor/i.test(c)
        );
        if (hasKnownKeywords) {
          headerRowIdx = r;
          break;
        }
        if (headerRowIdx === -1 && nonEmptyCells.length >= 2) {
          headerRowIdx = r;
        }
      }
    }

    if (headerRowIdx === -1) {
      headerRowIdx = 0;
    }

    const rawHeaders = (aoa[headerRowIdx] || []).map((h, colIdx) => {
      const cleanH = String(h ?? '').replace(/^\uFEFF/, '').trim();
      return cleanH || `عمود_${colIdx + 1}`;
    });

    const result: SheetRow[] = [];
    for (let r = headerRowIdx + 1; r < aoa.length; r++) {
      const row = aoa[r];
      if (!Array.isArray(row)) continue;
      const rowObj: SheetRow = {};
      let hasData = false;

      for (let c = 0; c < rawHeaders.length; c++) {
        const cellVal = row[c] !== undefined && row[c] !== null ? String(row[c]).trim() : '';
        const headerName = rawHeaders[c];
        rowObj[headerName] = cellVal;
        if (cellVal !== '') {
          hasData = true;
        }
      }

      if (hasData) {
        result.push(rowObj);
      }
    }

    return result;
  };

  // Check if workbook has specific Cash or Payment sheets
  const hasMultipleSheets = workbook.SheetNames.length > 1;
  let allRowsCombined: SheetRow[] = [];

  if (hasMultipleSheets) {
    let sheetIndicesToProcess: number[] = [];

    if (targetType === 'cash') {
      // Find cash sheets specifically (matching cash keywords or acceptor2/doner2)
      for (let sIdx = 0; sIdx < workbook.SheetNames.length; sIdx++) {
        const sheetName = workbook.SheetNames[sIdx];
        if (/كاش|cash|acceptor2|doner2|donor2|acceptor_2|doner_2|donor_2/i.test(sheetName)) {
          sheetIndicesToProcess.push(sIdx);
        }
      }
      // If no explicit match and 2+ sheets, tab 0 is standard cash template
      if (sheetIndicesToProcess.length === 0 && workbook.SheetNames.length >= 2) {
        sheetIndicesToProcess.push(0);
      }
    } else if (targetType === 'payments') {
      // Find payments sheets specifically
      for (let sIdx = 0; sIdx < workbook.SheetNames.length; sIdx++) {
        const sheetName = workbook.SheetNames[sIdx];
        if (/مدفوعات|payment|pay|acceptor1|doner1|donor1|acceptor(?![2-9])|doner(?![2-9])|donor(?![2-9])/i.test(sheetName)) {
          sheetIndicesToProcess.push(sIdx);
        }
      }
      // If no explicit match and 2+ sheets, tab 1 is standard payments template
      if (sheetIndicesToProcess.length === 0 && workbook.SheetNames.length >= 2) {
        sheetIndicesToProcess.push(1);
      }
    }

    // Default to all sheets if none filtered or targetType === 'machines'
    if (sheetIndicesToProcess.length === 0) {
      sheetIndicesToProcess = workbook.SheetNames.map((_, i) => i);
    }

    for (const sIdx of sheetIndicesToProcess) {
      const sheetName = workbook.SheetNames[sIdx];
      const ws = workbook.Sheets[sheetName];
      let candidateRows = extractRowsFromWorksheet(ws);
      if (!candidateRows.length) continue;

      let isCashSheet = /كاش|cash|acceptor2|doner2|donor2/i.test(sheetName);
      let isPaymentSheet = /مدفوعات|payment|pay|acceptor1|doner1|donor1/i.test(sheetName);

      // Support 2-tab template where Tab 1 (index 0) is Cash and Tab 2 (index 1) is Payments
      if (!isCashSheet && !isPaymentSheet && workbook.SheetNames.length >= 2) {
        if (sIdx === 0) isCashSheet = true;
        if (sIdx === 1) isPaymentSheet = true;
      }

      candidateRows = candidateRows.map((row) => ({
        ...row,
        _tabType: isCashSheet ? 'cash' : isPaymentSheet ? 'payments' : undefined,
        _tabName: sheetName,
      }));

      allRowsCombined = [...allRowsCombined, ...candidateRows];
    }
  }

  if (allRowsCombined.length > 0) {
    return filterRowsForTarget(allRowsCombined, targetType);
  }

  let bestRows: SheetRow[] = [];
  for (const sheetName of workbook.SheetNames) {
    const ws = workbook.Sheets[sheetName];
    const candidateRows = extractRowsFromWorksheet(ws);
    if (candidateRows.length > bestRows.length) {
      bestRows = candidateRows;
    }
  }

  if (bestRows.length === 0) {
    for (const sheetName of workbook.SheetNames) {
      const ws = workbook.Sheets[sheetName];
      const rawRows = XLSX.utils.sheet_to_json<SheetRow>(ws, { defval: '', raw: false });
      const cleaned = rawRows
        .map((row) => {
          const cRow: SheetRow = {};
          for (const key of Object.keys(row)) {
            const cleanKey = key.replace(/^\uFEFF/, '').trim();
            cRow[cleanKey] = typeof row[key] === 'string' ? row[key].trim() : row[key];
          }
          return cRow;
        })
        .filter((row) =>
          Object.values(row).some((v) => v !== undefined && v !== null && String(v).trim() !== '')
        );

      if (cleaned.length > bestRows.length) {
        bestRows = cleaned;
      }
    }
  }

  return filterRowsForTarget(bestRows, targetType);
}

export function downloadTemplateFile(templateType: 1 | 2 | 3, format: 'xlsx' | 'csv' = 'xlsx') {
  let filename = '';
  let headers: string[] = [];
  let sampleRows: (string | number)[][] = [];

  if (templateType === 1) {
    filename = 'نموذج_شيت_الماكينات_المستهدفة';
    headers = ['رقم الماكينة', 'الفرع / المنطقة', 'ملاحظات'];
    sampleRows = [
      ['POS-1001', 'فرع المعادي - القاهرة', 'ماكينة رئيسية'],
      ['POS-1002', 'فرع المهندسين - الجيزة', ''],
      ['POS-1003', 'فرع سموحة - الإسكندرية', ''],
    ];
  } else if (templateType === 2) {
    filename = 'نموذج_شيت_المدفوعات';
    headers = ['رقم الماكينة', 'رقم حساب المدفوعات', 'اسم مسؤول المدفوعات', 'حالة الحساب'];
    sampleRows = [
      ['POS-1001', '1234', 'أحمد محمود سالم', 'نشط'],
      ['POS-1001', '456', 'محمود حسن رضوان', 'نشط'],
    ];
  } else {
    filename = 'نموذج_شيت_الكاش';
    headers = ['رقم الماكينة', 'رقم حساب الكاش', 'اسم مسؤول الكاش', 'حالة الحساب'];
    sampleRows = [
      ['7-POS-1001', '1234', 'أحمد محمود سالم (كاش)', 'نشط'],
      ['7-POS-1002', '555', 'خالد عبد الرحمن (كاش)', 'نشط'],
    ];
  }

  const fullAoa = [headers, ...sampleRows];

  if (format === 'csv') {
    const csvContent =
      '\uFEFF' +
      fullAoa
        .map((row) =>
          row
            .map((val) => `"${(val ?? '').toString().replace(/"/g, '""')}"`)
            .join(',')
        )
        .join('\r\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${filename}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(link.href);
  } else {
    const wb = XLSX.utils.book_new();
    
    if (templateType === 1) {
      const cashSampleRows = [
        ['7-POS-1001', 'فرع المعادي - القاهرة', 'ماكينة كاش'],
        ['1234', 'فرع المهندسين - الجيزة', 'ماكينة كاش تحول لـ 7-1234 تلقائياً'],
      ];
      const ws1 = XLSX.utils.aoa_to_sheet([['رقم الماكينة', 'الفرع / المنطقة', 'ملاحظات'], ...cashSampleRows]);
      ws1['!cols'] = [{ wch: 22 }, { wch: 30 }, { wch: 25 }];
      XLSX.utils.book_append_sheet(wb, ws1, 'ماكينات الكاش');

      const ws2 = XLSX.utils.aoa_to_sheet([['رقم الماكينة', 'الفرع / المنطقة', 'ملاحظات'], ...sampleRows]);
      ws2['!cols'] = [{ wch: 22 }, { wch: 30 }, { wch: 25 }];
      XLSX.utils.book_append_sheet(wb, ws2, 'ماكينات المدفوعات');
    } else {
      const ws = XLSX.utils.aoa_to_sheet(fullAoa);
      ws['!cols'] = templateType === 2 ? [{ wch: 22 }, { wch: 24 }, { wch: 32 }, { wch: 18 }] : [{ wch: 22 }, { wch: 24 }, { wch: 32 }, { wch: 18 }];
      XLSX.utils.book_append_sheet(wb, ws, 'النموذج_المعتمد');
    }
    XLSX.writeFile(wb, `${filename}.xlsx`);
  }
}

export function exportReconciliationToExcel(rows: ExpandedRow[], fileName = 'تقرير_مطابقة_الماكينات_مدفوعات_وكاش.xlsx') {
  const wb = XLSX.utils.book_new();
  const headers = [
    'م',
    'رقم الماكينة',
    'نوع الشيت',
    'رقم الحساب',
    'اسم المسؤول / المندوب',
    'ترتيب السجل',
    'إجمالي السجلات',
    'حالة الماكينة',
  ];
  const colWidths = [
    { wch: 6 },
    { wch: 22 },
    { wch: 18 },
    { wch: 22 },
    { wch: 32 },
    { wch: 24 },
    { wch: 18 },
    { wch: 24 },
  ];

  const CHUNK_SIZE = 900000; // Under Excel's 1,048,576 limit

  if (rows.length <= CHUNK_SIZE) {
    const data: (string | number)[][] = [headers];

    for (let i = 0; i < rows.length; i++) {
      const item = rows[i];
      let st = 'شاغرة (بدون مسؤول)';
      if (item.status === 'single') st = 'مطابقة فردية';
      if (item.status === 'multi') st = `متعددة السجلات (${item.totalRepsForMachine})`;

      let orderLabel = 'غير مسند';
      if (item.status === 'single') orderLabel = 'وحيد (1 من 1)';
      if (item.status === 'multi') {
        orderLabel =
          item.repOrder === 1
            ? `الأول (1 من ${item.totalRepsForMachine})`
            : `إضافي (${item.repOrder} من ${item.totalRepsForMachine})`;
      }

      data.push([
        i + 1,
        item.machine,
        item.type === 'cash' ? '💵 كاش' : '💳 مدفوعات',
        item.account,
        item.repName,
        orderLabel,
        item.totalRepsForMachine,
        st,
      ]);
    }

    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = colWidths;
    XLSX.utils.book_append_sheet(wb, ws, 'مدفوعات_وكاش_الماكينات');
  } else {
    // Multi-sheet automatic splitting to preserve every row in Excel
    const numSheets = Math.ceil(rows.length / CHUNK_SIZE);
    for (let part = 0; part < numSheets; part++) {
      const start = part * CHUNK_SIZE;
      const end = Math.min(start + CHUNK_SIZE, rows.length);
      const slice = rows.slice(start, end);

      const data: (string | number)[][] = [headers];
      for (let i = 0; i < slice.length; i++) {
        const item = slice[i];
        let st = 'شاغرة (بدون مسؤول)';
        if (item.status === 'single') st = 'مطابقة فردية';
        if (item.status === 'multi') st = `متعددة السجلات (${item.totalRepsForMachine})`;

        let orderLabel = 'غير مسند';
        if (item.status === 'single') orderLabel = 'وحيد (1 من 1)';
        if (item.status === 'multi') {
          orderLabel =
            item.repOrder === 1
              ? `الأول (1 من ${item.totalRepsForMachine})`
              : `إضافي (${item.repOrder} من ${item.totalRepsForMachine})`;
        }

        data.push([
          start + i + 1,
          item.machine,
          item.type === 'cash' ? '💵 كاش' : '💳 مدفوعات',
          item.account,
          item.repName,
          orderLabel,
          item.totalRepsForMachine,
          st,
        ]);
      }

      const ws = XLSX.utils.aoa_to_sheet(data);
      ws['!cols'] = colWidths;
      XLSX.utils.book_append_sheet(wb, ws, `النتائج_جزء_${part + 1}_من_${numSheets}`);
    }
  }

  XLSX.writeFile(wb, fileName);
}

export function exportReconciliationToCsv(rows: ExpandedRow[], fileName = 'تقرير_مطابقة_الماكينات_شامل.csv') {
  const headers = [
    'م',
    'رقم الماكينة',
    'نوع الشيت',
    'رقم الحساب',
    'اسم المسؤول / المندوب',
    'ترتيب السجل',
    'إجمالي السجلات',
    'حالة الماكينة',
  ];

  const lines: string[] = [];
  lines.push(headers.map((h) => `"${h.replace(/"/g, '""')}"`).join(','));

  for (let i = 0; i < rows.length; i++) {
    const item = rows[i];
    let st = 'شاغرة (بدون مسؤول)';
    if (item.status === 'single') st = 'مطابقة فردية';
    if (item.status === 'multi') st = `متعددة السجلات (${item.totalRepsForMachine})`;

    let orderLabel = 'غير مسند';
    if (item.status === 'single') orderLabel = 'وحيد (1 من 1)';
    if (item.status === 'multi') {
      orderLabel =
        item.repOrder === 1
          ? `الأول (1 من ${item.totalRepsForMachine})`
          : `إضافي (${item.repOrder} من ${item.totalRepsForMachine})`;
    }

    const row = [
      i + 1,
      item.machine,
      item.type === 'cash' ? 'كاش' : 'مدفوعات',
      item.account,
      item.repName,
      orderLabel,
      item.totalRepsForMachine,
      st,
    ];

    lines.push(
      row
        .map((val) => {
          const str = String(val ?? '').replace(/"/g, '""');
          return `"${str}"`;
        })
        .join(',')
    );
  }

  const csvContent = '\uFEFF' + lines.join('\r\n');
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(link.href);
}
