import type { GanttTask } from './ganttTypes'

export type GanttSheetState = {
  id: string
  /** Tab label (like an Excel sheet name). */
  sheetName: string
  /** Used for XLSX export metadata. */
  projectName: string
  tasks: GanttTask[]
  viewRangeOverride: { start: string; end: string } | null
}

export type GanttWorkbookState = {
  activeSheetId: string
  sheets: GanttSheetState[]
}

export function createSheet(
  partial?: Partial<GanttSheetState> & { id?: string },
): GanttSheetState {
  return {
    id: partial?.id ?? crypto.randomUUID(),
    sheetName: partial?.sheetName ?? 'Sheet',
    projectName: partial?.projectName ?? '',
    tasks: partial?.tasks ?? [],
    viewRangeOverride: partial?.viewRangeOverride ?? null,
  }
}

export function nextSheetLabel(sheets: GanttSheetState[]): string {
  const used = new Set(sheets.map((s) => s.sheetName))
  let n = sheets.length + 1
  while (used.has(`Sheet ${n}`)) n += 1
  return `Sheet ${n}`
}
