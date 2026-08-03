import { pruneDanglingDeps } from './ganttDeps'
import type { GanttSheetState, GanttWorkbookState } from './ganttSheet'
import { createSheet } from './ganttSheet'
import type { GanttTask } from './ganttTypes'

/**
 * Validation for workbook payloads. Everything here treats its input as
 * untrusted — it parses localStorage written by an older build as well as rows
 * coming back over the network — so unknown shapes are dropped rather than
 * trusted.
 */

/** Current payload version. v2 (pre-dependencies) still parses. */
export const WORKBOOK_VERSION = 3
const SUPPORTED_VERSIONS = new Set([2, 3])

const ISO_DATE = /^\d{4}-\d{2}-\d{2}$/
const MAX_SHEET_NAME = 31

/** Dates are optional: anything that is not a well-formed ISO date is dropped. */
function parseOptionalDate(x: unknown): string | null {
  return typeof x === 'string' && ISO_DATE.test(x) ? x : null
}

function isGanttTask(x: unknown): x is Record<string, unknown> {
  if (!x || typeof x !== 'object') return false
  const o = x as Record<string, unknown>
  return (
    typeof o.id === 'string' &&
    typeof o.name === 'string' &&
    (typeof o.parentId === 'undefined' || typeof o.parentId === 'string') &&
    typeof o.progress === 'number' &&
    Number.isFinite(o.progress)
  )
}

function parseDeps(x: unknown): string[] | undefined {
  if (!Array.isArray(x)) return undefined
  const ids = [...new Set(x.filter((d): d is string => typeof d === 'string' && d.length > 0))]
  return ids.length ? ids : undefined
}

function normalizeTask(raw: Record<string, unknown>): GanttTask {
  const parentId = raw.parentId
  const progress = typeof raw.progress === 'number' ? raw.progress : 0
  return {
    id: raw.id as string,
    name: raw.name as string,
    parentId: typeof parentId === 'string' && parentId.length > 0 ? parentId : undefined,
    start: parseOptionalDate(raw.start),
    end: parseOptionalDate(raw.end),
    progress: Math.min(100, Math.max(0, Math.round(progress))),
    collapsed: raw.collapsed === true ? true : undefined,
    plotted: raw.plotted === false ? false : undefined,
    deps: parseDeps(raw.deps),
  }
}

function parseViewOverride(x: unknown): { start: string; end: string } | null {
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

export function parseSheet(o: unknown): GanttSheetState | null {
  if (!o || typeof o !== 'object') return null
  const r = o as Record<string, unknown>
  if (typeof r.id !== 'string' || typeof r.sheetName !== 'string') return null
  if (typeof r.projectName !== 'string') return null
  if (!Array.isArray(r.tasks)) return null
  const tasks = pruneDanglingDeps(r.tasks.filter(isGanttTask).map(normalizeTask))
  return {
    id: r.id,
    sheetName: r.sheetName.slice(0, MAX_SHEET_NAME),
    projectName: r.projectName,
    tasks,
    viewRangeOverride: parseViewOverride(r.viewRangeOverride),
  }
}

/** Parses an already-JSON-decoded workbook payload. Returns null if unusable. */
export function parseWorkbookPayload(data: unknown): GanttWorkbookState | null {
  if (!data || typeof data !== 'object') return null
  const o = data as Record<string, unknown>
  if (typeof o.v !== 'number' || !SUPPORTED_VERSIONS.has(o.v)) return null
  if (typeof o.activeSheetId !== 'string') return null
  if (!Array.isArray(o.sheets)) return null
  const sheets = o.sheets.map(parseSheet).filter((s): s is GanttSheetState => s !== null)
  if (!sheets.length) return null
  const activeOk = sheets.some((s) => s.id === o.activeSheetId)
  return {
    activeSheetId: activeOk ? o.activeSheetId : sheets[0].id,
    sheets,
  }
}

export function serializeWorkbook(state: GanttWorkbookState): {
  v: number
  activeSheetId: string
  sheets: GanttSheetState[]
} {
  return {
    v: WORKBOOK_VERSION,
    activeSheetId: state.activeSheetId,
    sheets: state.sheets,
  }
}

/** Migrates the original single-project payload into a one-sheet workbook. */
export function migrateLegacyV1(raw: string): GanttWorkbookState | null {
  try {
    const data = JSON.parse(raw) as unknown
    if (!data || typeof data !== 'object') return null
    const o = data as Record<string, unknown>
    if (o.v !== 1) return null
    if (typeof o.projectName !== 'string') return null
    if (!Array.isArray(o.tasks)) return null
    const sheet = createSheet({
      sheetName: 'Sheet 1',
      projectName: o.projectName,
      tasks: o.tasks.filter(isGanttTask).map(normalizeTask),
      viewRangeOverride: parseViewOverride(o.viewRangeOverride),
    })
    return { activeSheetId: sheet.id, sheets: [sheet] }
  } catch {
    return null
  }
}
