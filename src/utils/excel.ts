import * as XLSX from 'xlsx';
import { ExpandedRow, SheetRow } from '../types';

export async function parseExcelFile(file: File): Promise<SheetRow[]> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(new Uint8Array(buffer), { type: 'array' });
  const firstSheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheetName];
  return XLSX.utils.sheet_to_json<SheetRow>(worksheet, { defval: '' });
}

export function downloadTemplateFile(templateType: 1 | 2, format: 'xlsx' | 'csv' = 'xlsx') {
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
      ['POS-1004', 'فرع المشاية - المنصورة', ''],
      ['POS-1005', 'فرع الجامعة - طنطا', ''],
      ['POS-1006', 'فرع الجمهورية - أسيوط', ''],
    ];
  } else {
    filename = 'نموذج_شيت_ربط_المناديب_والحسابات';
    headers = ['رقم الماكينة', 'رقم حساب المندوب', 'اسم المندوب', 'حالة المندوب'];
    sampleRows = [
      ['POS-1001', '1234', 'أحمد محمود سالم', 'نشط'],
      ['POS-1001', '456', 'محمود حسن رضوان', 'نشط'],
      ['POS-1002', '555', 'خالد عبد الرحمن', 'نشط'],
      ['POS-1003', '789', 'طارق زياد العتيبي', 'نشط'],
      ['POS-1003', '890', 'عمر فاروق الشامي', 'نشط'],
      ['POS-1003', '999', 'إبراهيم حسني مراد', 'نشط'],
      ['POS-1004', '777', 'سامح عبد الله كمال', 'نشط'],
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
    const ws = XLSX.utils.aoa_to_sheet(fullAoa);
    if (templateType === 1) {
      ws['!cols'] = [{ wch: 22 }, { wch: 30 }, { wch: 25 }];
    } else {
      ws['!cols'] = [{ wch: 22 }, { wch: 24 }, { wch: 32 }, { wch: 18 }];
    }
    XLSX.utils.book_append_sheet(wb, ws, 'النموذج_المعتمد');
    XLSX.writeFile(wb, `${filename}.xlsx`);
  }
}

export function exportReconciliationToExcel(rows: ExpandedRow[], fileName = 'تقرير_مطابقة_الماكينات_أسطر_منفصلة_لكل_مندوب.xlsx') {
  const data: (string | number)[][] = [
    [
      'م',
      'رقم الماكينة',
      'رقم حساب المندوب',
      'اسم المندوب',
      'ترتيب المندوب للماكينة',
      'إجمالي مناديب الماكينة',
      'حالة الماكينة',
    ],
  ];

  for (let i = 0; i < rows.length; i++) {
    const item = rows[i];
    let st = 'شاغرة (بدون مندوب)';
    if (item.status === 'single') st = 'مندوب فردي';
    if (item.status === 'multi') st = `متعددة المناديب (${item.totalRepsForMachine} مناديب)`;

    let orderLabel = 'غير مسند';
    if (item.status === 'single') orderLabel = 'مندوب وحيد (1 من 1)';
    if (item.status === 'multi') {
      orderLabel =
        item.repOrder === 1
          ? `المندوب الأول (1 من ${item.totalRepsForMachine})`
          : `مندوب إضافي (${item.repOrder} من ${item.totalRepsForMachine})`;
    }

    data.push([
      i + 1,
      item.machine,
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
    { wch: 24 },
    { wch: 24 },
    { wch: 34 },
    { wch: 28 },
    { wch: 22 },
    { wch: 28 },
  ];

  XLSX.utils.book_append_sheet(wb, ws, 'أسطر_الماكينات_والمناديب');
  XLSX.writeFile(wb, fileName);
}
