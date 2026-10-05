export interface RepEntry {
  account: string;
  name: string;
  isMissingRepName: boolean;
}

export interface MachineSummary {
  index: number;
  machine: string;
  repCount: number;
  reps: RepEntry[];
  status: 'single' | 'multi' | 'none';
}

export interface ExpandedRow {
  id: string;
  machine: string;
  account: string;
  repName: string;
  repOrder: number;
  totalRepsForMachine: number;
  status: 'single' | 'multi' | 'none';
  machineGroupIndex: number;
  isFirstOfGroup: boolean;
  isLastOfGroup: boolean;
  isMissingRepName: boolean;
}

export interface SheetRow {
  [key: string]: string | number | undefined;
}

export interface DuplicateAuditItem {
  machine: string;
  occurrences: number;
}

export interface IrregularAccountItem {
  machine: string;
  account: string;
  repName: string;
  reason: string;
}
