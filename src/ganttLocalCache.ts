import type { GanttSheetState, GanttWorkbookState } from './ganttSheet'
import { createSheet } from './ganttSheet'
import type { GanttTask } from './ganttTypes'

const STORAGE_KEY = 'gantt-chart-helper'
/** Legacy payload shape (single project). */
const LEGACY_STORAGE_KEY = 'gantt-chart-helper:v1'

const ISO_DATE = /^\d{4}-\d{2}-\d{2}$/

function isGanttTask(x: unknown): x is GanttTask {
  if (!x || typeof x !== 'object') return false
  const o = x as Record<string, unknown>
  return (
    typeof o.id === 'string' &&
    typeof o.name === 'string' &&
    typeof o.start === 'string' &&
    typeof o.end === 'string' &&
    ISO_DATE.test(o.start) &&
    ISO_DATE.test(o.end) &&
    typeof o.progress === 'number' &&
    Number.isFinite(o.progress)
  )
}

function normalizeTask(t: GanttTask): GanttTask {
  return {
    ...t,
    progress: Math.min(100, Math.max(0, Math.round(t.progress))),
  }
}

function parseViewOverride(
  x: unknown,
): { start: string; end: string } | null {
  if (x === null) return null
  if (!x || typeof x !== 'object') return null
  const o = x as Record<string, unknown>
  if (
    typeof o.start !== 'string' ||
    typeof o.end !== 'string' ||
    !ISO_DATE.test(o.start) ||
    !ISO_DATE.test(o.end)
  ) {
    return null
  }
  return { start: o.start, end: o.end }
}

function parseSheet(o: unknown): GanttSheetState | null {
  if (!o || typeof o !== 'object') return null
  const r = o as Record<string, unknown>
  if (typeof r.id !== 'string' || typeof r.sheetName !== 'string') return null
  if (typeof r.projectName !== 'string') return null
  if (!Array.isArray(r.tasks)) return null
  const tasks = r.tasks.filter(isGanttTask).map(normalizeTask)
  const viewRangeOverride = parseViewOverride(r.viewRangeOverride)
  return {
    id: r.id,
    sheetName: r.sheetName.slice(0, 31),
    projectName: r.projectName,
    tasks,
    viewRangeOverride,
  }
}

function migrateLegacyV1(raw: string): GanttWorkbookState | null {
  try {
    const data = JSON.parse(raw) as unknown
    if (!data || typeof data !== 'object') return null
    const o = data as Record<string, unknown>
    if (o.v !== 1) return null
    if (typeof o.projectName !== 'string') return null
    if (!Array.isArray(o.tasks)) return null
    const tasks = o.tasks.filter(isGanttTask).map(normalizeTask)
    const viewRangeOverride = parseViewOverride(o.viewRangeOverride)
    const sheet = createSheet({
      sheetName: 'Sheet 1',
      projectName: o.projectName,
      tasks,
      viewRangeOverride,
    })
    return { activeSheetId: sheet.id, sheets: [sheet] }
  } catch {
    return null
  }
}

export function loadGanttWorkbook(): GanttWorkbookState | null {
  if (typeof localStorage === 'undefined') return null
  try {
    const raw = localStorage.getItem(STORAGE_KEY)
    if (raw) {
      const data = JSON.parse(raw) as unknown
      if (!data || typeof data !== 'object') return null
      const o = data as Record<string, unknown>
      if (o.v !== 2) return null
      if (typeof o.activeSheetId !== 'string') return null
      if (!Array.isArray(o.sheets)) return null
      const sheets = o.sheets
        .map(parseSheet)
        .filter((s): s is GanttSheetState => s !== null)
      if (!sheets.length) return null
      const activeOk = sheets.some((s) => s.id === o.activeSheetId)
      return {
        activeSheetId: activeOk ? o.activeSheetId : sheets[0].id,
        sheets,
      }
    }

    const legacy = localStorage.getItem(LEGACY_STORAGE_KEY)
    if (legacy) {
      const migrated = migrateLegacyV1(legacy)
      if (migrated) {
        try {
          localStorage.removeItem(LEGACY_STORAGE_KEY)
        } catch {
          /* ignore */
        }
        return migrated
      }
    }
  } catch {
    return null
  }
  return null
}

export function saveGanttWorkbook(state: GanttWorkbookState): void {
  if (typeof localStorage === 'undefined') return
  try {
    const payload = JSON.stringify({
      v: 2,
      activeSheetId: state.activeSheetId,
      sheets: state.sheets,
    })
    localStorage.setItem(STORAGE_KEY, payload)
  } catch {
    // quota / private mode
  }
}
