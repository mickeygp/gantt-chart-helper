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

/** Excel caps sheet names at 31 characters, and so does the tab rename input. */
const MAX_SHEET_NAME = 31

/**
 * "Plan A" → "Plan A copy" → "Plan A copy 2", trimmed to fit the name limit
 * without ever colliding with an existing tab.
 */
export function copySheetLabel(base: string, sheets: GanttSheetState[]): string {
  const used = new Set(sheets.map((s) => s.sheetName))
  const stem = base.replace(/ copy(?: \d+)?$/, '').trim() || 'Sheet'
  for (let n = 1; n < 1000; n += 1) {
    const suffix = n === 1 ? ' copy' : ` copy ${n}`
    const room = MAX_SHEET_NAME - suffix.length
    const candidate = `${stem.slice(0, room)}${suffix}`
    if (!used.has(candidate)) return candidate
  }
  return crypto.randomUUID().slice(0, MAX_SHEET_NAME)
}

/**
 * Deep-copies a sheet under fresh ids. Task ids are remapped so the copy's
 * parent links and dependencies point inside the copy — otherwise editing
 * "Plan B" would reach back into "Plan A".
 */
export function duplicateSheet(
  sheet: GanttSheetState,
  sheets: GanttSheetState[],
): GanttSheetState {
  const idMap = new Map<string, string>(
    sheet.tasks.map((t) => [t.id, crypto.randomUUID()]),
  )
  const tasks = sheet.tasks.map((t) => {
    const deps = t.deps
      ?.map((d) => idMap.get(d))
      .filter((d): d is string => d !== undefined)
    return {
      ...t,
      id: idMap.get(t.id)!,
      parentId: t.parentId ? idMap.get(t.parentId) : undefined,
      deps: deps?.length ? deps : undefined,
    }
  })
  return {
    id: crypto.randomUUID(),
    sheetName: copySheetLabel(sheet.sheetName, sheets),
    projectName: sheet.projectName,
    tasks,
    viewRangeOverride: sheet.viewRangeOverride
      ? { ...sheet.viewRangeOverride }
      : null,
  }
}
