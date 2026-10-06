import * as XLSX from 'xlsx';
import { ExpandedRow, SheetRow } from '../types';

export async function parseExcelFile(file: File): Promise<SheetRow[]> {
  const buffer = await file.arrayBuffer();
  let workbook: XLSX.WorkBook | null = null;

  try {
    workbook = XLSX.read(new Uint8Array(buffer), { type: 'array', codepage: 65001 });
  } catch (err1) {
    try {
      const text = new TextDecoder('utf-8').decode(buffer);
      workbook = XLSX.read(text, { type: 'string' });
    } catch (err2) {
      try {
        const text1256 = new TextDecoder('windows-1256').decode(buffer);
        workbook = XLSX.read(text1256, { type: 'string' });
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
    const aoa = XLSX.utils.sheet_to_json<any[]>(worksheet, { header: 1, defval: '', raw: false });
    if (!aoa || aoa.length === 0) return [];

    let headerRowIdx = -1;
    for (let r = 0; r < Math.min(aoa.length, 12); r++) {
      const row = aoa[r];
      if (!Array.isArray(row)) continue;
      const nonEmptyCells = row.map((c) => String(c ?? '').trim()).filter((c) => c !== '');
      if (nonEmptyCells.length >= 1) {
        const hasKnownKeywords = nonEmptyCells.some((c) =>
          /ماكينة|جهاز|pos|machine|sn|id|كود|حساب|مندوب|اسم|فرع|تاريخ|مسلسل|code|acc|name|terminal|merchant/i.test(c)
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
    for (const sheetName of workbook.SheetNames) {
      const ws = workbook.Sheets[sheetName];
      let candidateRows = extractRowsFromWorksheet(ws);
      if (!candidateRows.length) continue;

      const isCashSheet = /كاش|cash/i.test(sheetName);
      const isPaymentSheet = /مدفوعات|payment/i.test(sheetName);

      candidateRows = candidateRows.map((row) => ({
        ...row,
        _tabType: isCashSheet ? 'cash' : isPaymentSheet ? 'payments' : undefined,
      }));

      allRowsCombined = [...allRowsCombined, ...candidateRows];
    }
  }

  if (allRowsCombined.length > 0) {
    return allRowsCombined;
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

  return bestRows;
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
  const data: (string | number)[][] = [
    [
      'م',
      'رقم الماكينة',
      'نوع الشيت',
      'رقم الحساب',
      'اسم المسؤول / المندوب',
      'ترتيب السجل',
      'إجمالي السجلات',
      'حالة الماكينة',
    ],
  ];

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

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(data);

  ws['!cols'] = [
    { wch: 6 },
    { wch: 22 },
    { wch: 18 },
    { wch: 22 },
    { wch: 32 },
    { wch: 24 },
    { wch: 18 },
    { wch: 24 },
  ];

  XLSX.utils.book_append_sheet(wb, ws, 'مدفوعات_وكاش_الماكينات');
  XLSX.writeFile(wb, fileName);
}
