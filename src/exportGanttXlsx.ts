import * as XLSX from 'xlsx-js-style'

import {
  getEffectiveProgress,
  getVisibleTasks,
  isPlotted,
  resolveTaskRanges,
  unionRanges,
  type GanttTask,
  type TaskRange,
} from './ganttTypes'
import { addDaysISO, daysInclusive, parseISOToUtcMs } from './ganttDates'
import { validDeps } from './ganttDeps'
import {
  buildMonthSpans,
  buildWeekSpans,
  formatMonthYear,
  type WeekLabelFormat,
} from './ganttTimeline'

export type ExportGanttOptions = {
  projectName: string
  visibleRange?: { start: string; end: string }
  includeDayColumns?: boolean
  weekLabelFormat?: WeekLabelFormat
  /** Used for filename and workbook metadata; defaults to `new Date()`. */
  exportedAt?: Date
}

// ── Brand palette (matches in-app UI) ──────────────────────
const BRAND = {
  primary: '6366F1',
  primaryDark: '4338CA',
  accentBg: 'EEF2FF',
  textDark: '0F172A',
  textMuted: '64748B',
  textLight: '94A3B8',
  border: 'D1D5DB',
  borderLight: 'E5E7EB',
  white: 'FFFFFF',
  altRow: 'F8FAFC',
  weekBg: 'F1F5F9',
  monthBg: 'E0E7FF',
  todayCol: 'FEE2E2',
  successGreen: '10B981',
  warnAmber: 'F59E0B',
  dangerRed: 'EF4444',
} as const

// 8-color palette mirroring the UI. Each row gets a distinct hue.
const TASK_BAR_PALETTE: readonly { done: string; partial: string; planned: string }[] = [
  { done: '6366F1', partial: 'A5B4FC', planned: 'E0E7FF' }, // indigo
  { done: '0EA5E9', partial: '7DD3FC', planned: 'E0F2FE' }, // sky
  { done: '10B981', partial: '6EE7B7', planned: 'D1FAE5' }, // emerald
  { done: 'F59E0B', partial: 'FCD34D', planned: 'FEF3C7' }, // amber
  { done: 'EF4444', partial: 'FCA5A5', planned: 'FEE2E2' }, // red
  { done: 'EC4899', partial: 'F9A8D4', planned: 'FCE7F3' }, // pink
  { done: '8B5CF6', partial: 'C4B5FD', planned: 'EDE9FE' }, // violet
  { done: '14B8A6', partial: '5EEAD4', planned: 'CCFBF1' }, // teal
]

function paletteFor(idx: number) {
  return TASK_BAR_PALETTE[idx % TASK_BAR_PALETTE.length]!
}

// ── Style primitives ───────────────────────────────────────
type CellStyle = Record<string, unknown>

const thin = { style: 'thin', color: { rgb: BRAND.borderLight } }
const thinDark = { style: 'thin', color: { rgb: BRAND.border } }
const allBorders = { top: thin, bottom: thin, left: thin, right: thin }
const allBordersDark = { top: thinDark, bottom: thinDark, left: thinDark, right: thinDark }

const styleTitle: CellStyle = {
  font: { name: 'Calibri', sz: 20, bold: true, color: { rgb: BRAND.textDark } },
  alignment: { vertical: 'center', horizontal: 'left' },
}

const styleSubtitle: CellStyle = {
  font: { name: 'Calibri', sz: 11, color: { rgb: BRAND.textMuted } },
  alignment: { vertical: 'center', horizontal: 'left' },
}

const styleSectionLabel: CellStyle = {
  font: { name: 'Calibri', sz: 10, bold: true, color: { rgb: BRAND.textLight } },
  alignment: { vertical: 'center', horizontal: 'left' },
}

const styleTableHeader: CellStyle = {
  font: { name: 'Calibri', sz: 11, bold: true, color: { rgb: BRAND.white } },
  fill: { patternType: 'solid', fgColor: { rgb: BRAND.primary } },
  alignment: { vertical: 'center', horizontal: 'left' },
  border: allBordersDark,
}

const styleTableHeaderCenter: CellStyle = {
  ...styleTableHeader,
  alignment: { vertical: 'center', horizontal: 'center' },
}

const styleBody: CellStyle = {
  font: { name: 'Calibri', sz: 11, color: { rgb: BRAND.textDark } },
  alignment: { vertical: 'center', horizontal: 'left', wrapText: false },
  border: allBorders,
}

const styleBodyAlt: CellStyle = {
  ...styleBody,
  fill: { patternType: 'solid', fgColor: { rgb: BRAND.altRow } },
}

const styleBodyMuted: CellStyle = {
  ...styleBody,
  font: { name: 'Calibri', sz: 11, color: { rgb: BRAND.textMuted } },
}

function bodyCenter(alt: boolean): CellStyle {
  return {
    ...(alt ? styleBodyAlt : styleBody),
    alignment: { vertical: 'center', horizontal: 'center' },
  }
}

function bodyLeft(alt: boolean, bold = false): CellStyle {
  const base = alt ? styleBodyAlt : styleBody
  if (!bold) return base
  return {
    ...base,
    font: { name: 'Calibri', sz: 11, bold: true, color: { rgb: BRAND.textDark } },
  }
}

/** Greyed-italic date: the value was rolled up from subtasks, not entered. */
function dateFallbackStyle(alt: boolean): CellStyle {
  return {
    ...(alt ? styleBodyAlt : styleBody),
    font: { name: 'Calibri', sz: 10, italic: true, color: { rgb: BRAND.textLight } },
    alignment: { vertical: 'center', horizontal: 'center' },
  }
}

function progressStyle(alt: boolean, progress: number): CellStyle {
  const fillRgb =
    progress >= 100
      ? BRAND.successGreen
      : progress > 0
        ? BRAND.warnAmber
        : BRAND.textLight
  return {
    ...(alt ? styleBodyAlt : styleBody),
    font: { name: 'Calibri', sz: 11, bold: progress === 100, color: { rgb: fillRgb } },
    alignment: { vertical: 'center', horizontal: 'center' },
    numFmt: '0"%"',
  }
}

function ganttHeaderMonth(): CellStyle {
  return {
    font: { name: 'Calibri', sz: 11, bold: true, color: { rgb: BRAND.primaryDark } },
    fill: { patternType: 'solid', fgColor: { rgb: BRAND.monthBg } },
    alignment: { vertical: 'center', horizontal: 'center' },
    border: allBordersDark,
  }
}

function ganttHeaderWeek(): CellStyle {
  return {
    font: { name: 'Calibri', sz: 9, color: { rgb: BRAND.textMuted } },
    fill: { patternType: 'solid', fgColor: { rgb: BRAND.weekBg } },
    alignment: { vertical: 'center', horizontal: 'center' },
    border: allBorders,
  }
}

function ganttHeaderDay(isToday: boolean): CellStyle {
  return {
    font: {
      name: 'Calibri',
      sz: 8,
      bold: isToday,
      color: { rgb: isToday ? BRAND.dangerRed : BRAND.textMuted },
    },
    fill: { patternType: 'solid', fgColor: { rgb: isToday ? BRAND.todayCol : BRAND.altRow } },
    alignment: { vertical: 'center', horizontal: 'center' },
    border: allBorders,
  }
}

function ganttTaskCellEmpty(alt: boolean, isTodayCol: boolean): CellStyle {
  const baseFill = isTodayCol ? BRAND.todayCol : alt ? BRAND.altRow : BRAND.white
  return {
    fill: { patternType: 'solid', fgColor: { rgb: baseFill } },
    border: allBorders,
  }
}

function ganttTaskCellFilled(rgb: string): CellStyle {
  return {
    fill: { patternType: 'solid', fgColor: { rgb } },
    border: allBorders,
  }
}

function ganttTaskNameCell(alt: boolean, color: string, depth = 0): CellStyle {
  const isSubtask = depth > 0
  return {
    font: {
      name: 'Calibri',
      sz: isSubtask ? 10 : 11,
      bold: !isSubtask,
      italic: isSubtask,
      color: { rgb: isSubtask ? BRAND.textMuted : BRAND.textDark },
    },
    fill: { patternType: 'solid', fgColor: { rgb: alt ? BRAND.altRow : BRAND.white } },
    alignment: { vertical: 'center', horizontal: 'left', indent: 1 + depth * 2 },
    border: {
      top: thin,
      bottom: thin,
      right: thinDark,
      left: { style: isSubtask ? 'thin' : 'thick', color: { rgb: color } },
    },
  }
}

// ── Cell helpers ───────────────────────────────────────────
type Primitive = string | number

function setCell(
  ws: XLSX.WorkSheet,
  r: number,
  c: number,
  value: Primitive | null,
  style?: CellStyle,
): void {
  const addr = XLSX.utils.encode_cell({ r, c })
  const cell: XLSX.CellObject = {
    t: typeof value === 'number' ? 'n' : 's',
    v: value ?? '',
  }
  if (style) cell.s = style
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  ;(ws as any)[addr] = cell
}

function applyRange(
  ws: XLSX.WorkSheet,
  r0: number,
  c0: number,
  r1: number,
  c1: number,
): void {
  const refRange = ws['!ref']
    ? XLSX.utils.decode_range(ws['!ref'])
    : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } }
  refRange.s.r = Math.min(refRange.s.r, r0)
  refRange.s.c = Math.min(refRange.s.c, c0)
  refRange.e.r = Math.max(refRange.e.r, r1)
  refRange.e.c = Math.max(refRange.e.c, c1)
  ws['!ref'] = XLSX.utils.encode_range(refRange)
}

function exportDateYYYYMMDD(d: Date): string {
  const y = d.getFullYear()
  const m = String(d.getMonth() + 1).padStart(2, '0')
  const day = String(d.getDate()).padStart(2, '0')
  return `${y}-${m}-${day}`
}

const FILENAME_SPECIAL_RE = /[<>:"/\\|?*]/g

/** Safe file base name segment (Windows/macOS/Linux). Empty input becomes Untitled. */
export function sanitizeProjectNameForFilename(name: string): string {
  const trimmed = name.trim()
  const cleaned = [...trimmed]
    .map((ch) => {
      const code = ch.codePointAt(0) ?? 0
      if (code <= 31) return '_'
      return ch
    })
    .join('')
    .replace(FILENAME_SPECIAL_RE, '_')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 120)
  return cleaned.length ? cleaned : 'Untitled'
}

export function buildGanttExportBasename(projectName: string, exportedAt = new Date()): string {
  const safe = sanitizeProjectNameForFilename(projectName)
  const date = exportDateYYYYMMDD(exportedAt)
  return `${safe}_${date}`
}

function clampRangeOrder(start: string, end: string): { start: string; end: string } {
  if (parseISOToUtcMs(end) < parseISOToUtcMs(start)) return { start, end: start }
  return { start, end }
}

function iterateDays(fromISO: string, toISO: string): string[] {
  const out: string[] = []
  let cur = fromISO
  while (parseISOToUtcMs(cur) <= parseISOToUtcMs(toISO)) {
    out.push(cur)
    cur = addDaysISO(cur, 1)
  }
  return out
}

function buildDepthMap(tasks: GanttTask[]): Map<string, number> {
  const byId = new Map(tasks.map((t) => [t.id, t]))
  const map = new Map<string, number>()
  for (const t of tasks) {
    let depth = 0
    let cur = t.parentId
    while (cur) {
      const parent = byId.get(cur)
      if (!parent) break
      depth += 1
      if (depth > 12) break
      cur = parent.parentId
    }
    map.set(t.id, depth)
  }
  return map
}

function summarizeTasks(tasks: GanttTask[], ranges: Map<string, TaskRange | null>) {
  if (!tasks.length) {
    return { totalDays: 0, avgProgress: 0, completed: 0, unscheduled: 0 }
  }
  const totalDays = tasks.reduce((sum, t) => {
    const r = ranges.get(t.id)
    return sum + (r ? Math.max(0, daysInclusive(r.start, r.end)) : 0)
  }, 0)
  const avgProgress = Math.round(
    tasks.reduce((sum, t) => sum + getEffectiveProgress(t, tasks), 0) / tasks.length,
  )
  const completed = tasks.filter((t) => getEffectiveProgress(t, tasks) >= 100).length
  const unscheduled = tasks.filter((t) => !ranges.get(t.id)).length
  return { totalDays, avgProgress, completed, unscheduled }
}

// ── Project sheet ──────────────────────────────────────────
function buildProjectSheet(
  projectName: string,
  exportDate: string,
  tasks: GanttTask[],
  plottedTasks: GanttTask[],
  ranges: Map<string, TaskRange | null>,
  range: { start: string; end: string },
): XLSX.WorkSheet {
  const ws: XLSX.WorkSheet = {}
  const summary = summarizeTasks(tasks, ranges)

  setCell(ws, 0, 0, projectName, styleTitle)
  setCell(
    ws,
    1,
    0,
    `Exported ${exportDate} · ${tasks.length} task${tasks.length === 1 ? '' : 's'}`,
    styleSubtitle,
  )

  setCell(ws, 3, 0, 'PROJECT DETAILS', styleSectionLabel)
  setCell(ws, 4, 0, 'Field', styleTableHeader)
  setCell(ws, 4, 1, 'Value', styleTableHeader)

  const detailRows: [string, string | number, boolean][] = [
    ['Project name', projectName, false],
    ['Export date', exportDate, true],
    ['Visible range', `${range.start} → ${range.end}`, false],
    ['Tasks exported', tasks.length, true],
    ['Plotted on timeline', `${plottedTasks.length} / ${tasks.length}`, false],
  ]
  detailRows.forEach(([label, value, alt], i) => {
    const r = 5 + i
    setCell(ws, r, 0, label, bodyLeft(alt, true))
    setCell(ws, r, 1, value, bodyLeft(alt))
  })

  const summaryStart = 5 + detailRows.length + 1
  setCell(ws, summaryStart, 0, 'SUMMARY', styleSectionLabel)
  setCell(ws, summaryStart + 1, 0, 'Metric', styleTableHeader)
  setCell(ws, summaryStart + 1, 1, 'Value', styleTableHeader)

  const summaryRows: [string, string | number, boolean][] = [
    ['Total task-days', summary.totalDays, false],
    ['Average progress', `${summary.avgProgress}%`, true],
    ['Completed tasks', `${summary.completed} / ${tasks.length}`, false],
    ['Tasks without dates', summary.unscheduled, true],
  ]
  summaryRows.forEach(([label, value, alt], i) => {
    const r = summaryStart + 2 + i
    setCell(ws, r, 0, label, bodyLeft(alt, true))
    setCell(ws, r, 1, value, bodyLeft(alt))
  })

  const lastRow = summaryStart + 1 + summaryRows.length
  applyRange(ws, 0, 0, lastRow, 1)
  ws['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 1 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 1 } },
    { s: { r: 3, c: 0 }, e: { r: 3, c: 1 } },
    { s: { r: summaryStart, c: 0 }, e: { r: summaryStart, c: 1 } },
  ]
  ws['!cols'] = [{ wch: 22 }, { wch: 50 }]
  const rowHeights: { hpx: number }[] = []
  rowHeights[0] = { hpx: 32 }
  rowHeights[1] = { hpx: 22 }
  rowHeights[3] = { hpx: 22 }
  rowHeights[summaryStart] = { hpx: 22 }
  ws['!rows'] = rowHeights

  return ws
}

// ── Tasks sheet ────────────────────────────────────────────
function buildTasksSheet(
  projectName: string,
  tasks: GanttTask[],
  ranges: Map<string, TaskRange | null>,
): XLSX.WorkSheet {
  const ws: XLSX.WorkSheet = {}
  const headers = [
    'Task',
    'Start',
    'End',
    'Duration (days)',
    'Progress',
    'Plotted',
    'Depends on',
  ] as const
  const nameById = new Map(tasks.map((t) => [t.id, t.name.trim() || '(untitled)']))

  setCell(ws, 0, 0, `Tasks — ${projectName}`, styleTitle)
  setCell(
    ws,
    1,
    0,
    `${tasks.length} task${tasks.length === 1 ? '' : 's'}`,
    styleSubtitle,
  )

  const headerRow = 3
  headers.forEach((h, c) => {
    setCell(ws, headerRow, c, h, c === 0 ? styleTableHeader : styleTableHeaderCenter)
  })

  const depthMap = buildDepthMap(tasks)

  if (!tasks.length) {
    const r = headerRow + 1
    setCell(ws, r, 0, '(no tasks to export)', styleBodyMuted)
    for (let c = 1; c < headers.length; c += 1) setCell(ws, r, c, '', styleBody)
    applyRange(ws, 0, 0, r, headers.length - 1)
  } else {
    tasks.forEach((t, i) => {
      const r = headerRow + 1 + i
      const alt = i % 2 === 1
      const depth = depthMap.get(t.id) ?? 0
      const isSubtask = depth > 0
      const color = paletteFor(i).done
      const namePrefix = isSubtask ? '↳ ' : ''
      setCell(ws, r, 0, `${namePrefix}${t.name.trim() || '(untitled)'}`, {
        font: {
          name: 'Calibri',
          sz: isSubtask ? 10 : 11,
          bold: !isSubtask,
          italic: isSubtask,
          color: { rgb: isSubtask ? BRAND.textMuted : BRAND.textDark },
        },
        fill: { patternType: 'solid', fgColor: { rgb: alt ? BRAND.altRow : BRAND.white } },
        alignment: { vertical: 'center', horizontal: 'left', indent: 1 + depth * 2 },
        border: {
          top: thin,
          bottom: thin,
          right: thin,
          left: { style: isSubtask ? 'thin' : 'thick', color: { rgb: color } },
        },
      })
      // Dates are optional; a blank field falls back to the subtask roll-up,
      // which is shown greyed out so it reads as inherited rather than entered.
      const range = ranges.get(t.id) ?? null
      const startStyle = t.start === null ? dateFallbackStyle(alt) : bodyCenter(alt)
      const endStyle = t.end === null ? dateFallbackStyle(alt) : bodyCenter(alt)
      setCell(ws, r, 1, t.start ?? range?.start ?? '—', startStyle)
      setCell(ws, r, 2, t.end ?? range?.end ?? '—', endStyle)
      setCell(
        ws,
        r,
        3,
        range ? Math.max(0, daysInclusive(range.start, range.end)) : '—',
        bodyCenter(alt),
      )
      const pct = getEffectiveProgress(t, tasks)
      setCell(
        ws,
        r,
        4,
        Math.min(100, Math.max(0, Math.round(pct))),
        progressStyle(alt, pct),
      )
      setCell(ws, r, 5, isPlotted(t) ? 'Yes' : 'No', {
        ...bodyCenter(alt),
        font: {
          name: 'Calibri',
          sz: 11,
          color: { rgb: isPlotted(t) ? BRAND.textDark : BRAND.textLight },
        },
      })
      // Predecessors by name — ids mean nothing to whoever opens the file.
      const depNames = validDeps(t, tasks)
        .map((id) => nameById.get(id))
        .filter((n): n is string => n !== undefined)
      setCell(ws, r, 6, depNames.length ? depNames.join(', ') : '—', {
        ...bodyCenter(alt),
        alignment: { vertical: 'center', horizontal: 'left', indent: 1 },
        font: {
          name: 'Calibri',
          sz: 10,
          color: { rgb: depNames.length ? BRAND.textDark : BRAND.textLight },
        },
      })
    })
    applyRange(ws, 0, 0, headerRow + tasks.length, headers.length - 1)
  }

  ws['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: headers.length - 1 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: headers.length - 1 } },
  ]
  ws['!cols'] = [
    { wch: 38 },
    { wch: 13 },
    { wch: 13 },
    { wch: 16 },
    { wch: 12 },
    { wch: 10 },
    { wch: 28 },
  ]
  const rowHeights: { hpx: number }[] = []
  rowHeights[0] = { hpx: 32 }
  rowHeights[1] = { hpx: 22 }
  rowHeights[headerRow] = { hpx: 24 }
  ws['!rows'] = rowHeights

  return ws
}

// ── Gantt sheet ────────────────────────────────────────────
function buildGanttSheet(
  projectName: string,
  allTasks: GanttTask[],
  /** Only the rows the app plots: expanded and ticked, in the same order. */
  rows: GanttTask[],
  ranges: Map<string, TaskRange | null>,
  visibleRange: { start: string; end: string },
  includeDayColumns: boolean,
  exportDate: string,
  weekLabelFormat: WeekLabelFormat,
): XLSX.WorkSheet {
  const ws: XLSX.WorkSheet = {}
  const days = iterateDays(visibleRange.start, visibleRange.end)
  const monthSpans = buildMonthSpans(days)
  const weekSpans = buildWeekSpans(days, weekLabelFormat)
  const todayIdx = days.indexOf(exportDate)

  const weekBuckets = (() => {
    const buckets: { label: string; dayCount: number; startOffset: number; endOffset: number }[] =
      []
    let startOffset = 0
    for (const span of weekSpans) {
      const endOffset = startOffset + span.dayCount - 1
      buckets.push({ label: span.label, dayCount: span.dayCount, startOffset, endOffset })
      startOffset = endOffset + 1
    }
    return buckets
  })()

  const cols = 1 + (includeDayColumns ? days.length : weekSpans.length)

  setCell(ws, 0, 0, `Timeline — ${projectName}`, styleTitle)
  setCell(
    ws,
    1,
    0,
    `${visibleRange.start} → ${visibleRange.end} · ${days.length} day${
      days.length === 1 ? '' : 's'
    }`,
    styleSubtitle,
  )

  const monthHeaderRow = 3
  const weekHeaderRow = 4
  const dayHeaderRow = includeDayColumns ? 5 : -1
  const taskStartRow = includeDayColumns ? 6 : 5

  // Task column header (merged across header rows)
  const taskHeaderTopRow = monthHeaderRow
  const taskHeaderBottomRow = includeDayColumns ? dayHeaderRow : weekHeaderRow
  setCell(ws, taskHeaderTopRow, 0, 'Task', {
    ...styleTableHeader,
    alignment: { vertical: 'center', horizontal: 'left', indent: 1 },
  })
  for (let r = taskHeaderTopRow + 1; r <= taskHeaderBottomRow; r += 1) {
    setCell(ws, r, 0, '', styleTableHeader)
  }

  // Month header row
  if (includeDayColumns) {
    let monthStartCol = 1
    for (const span of monthSpans) {
      setCell(ws, monthHeaderRow, monthStartCol, span.label, ganttHeaderMonth())
      for (let i = 1; i < span.dayCount; i += 1) {
        setCell(ws, monthHeaderRow, monthStartCol + i, '', ganttHeaderMonth())
      }
      monthStartCol += span.dayCount
    }
  } else {
    let weekCol = 1
    while (weekCol <= weekBuckets.length) {
      const firstBucket = weekBuckets[weekCol - 1]!
      const firstDay = days[firstBucket.startOffset] ?? visibleRange.start
      const monthKey = firstDay.slice(0, 7)
      const startCol = weekCol
      let endCol = startCol
      while (endCol < weekBuckets.length) {
        const nextBucket = weekBuckets[endCol]!
        const nextFirstDay = days[nextBucket.startOffset] ?? visibleRange.end
        if (nextFirstDay.slice(0, 7) !== monthKey) break
        endCol += 1
      }
      setCell(
        ws,
        monthHeaderRow,
        startCol,
        formatMonthYear(`${monthKey}-01`),
        ganttHeaderMonth(),
      )
      for (let c = startCol + 1; c <= endCol; c += 1) {
        setCell(ws, monthHeaderRow, c, '', ganttHeaderMonth())
      }
      weekCol = endCol + 1
    }
  }

  // Week header row
  let weekStartCol = 1
  for (const bucket of weekBuckets) {
    setCell(ws, weekHeaderRow, weekStartCol, bucket.label, ganttHeaderWeek())
    if (includeDayColumns) {
      for (let i = 1; i < bucket.dayCount; i += 1) {
        setCell(ws, weekHeaderRow, weekStartCol + i, '', ganttHeaderWeek())
      }
      weekStartCol += bucket.dayCount
    } else {
      weekStartCol += 1
    }
  }

  // Day header row
  if (includeDayColumns) {
    days.forEach((d, idx) => {
      setCell(
        ws,
        dayHeaderRow,
        1 + idx,
        Number(d.slice(8, 10)),
        ganttHeaderDay(idx === todayIdx),
      )
    })
  }

  // Task rows
  const rangeStartMs = parseISOToUtcMs(visibleRange.start)
  const rangeEndMs = parseISOToUtcMs(visibleRange.end)
  const dayMs = 86_400_000
  const depthMap = buildDepthMap(allTasks)
  // Colors follow the position in the full task list so they match the app.
  const colorIndexById = new Map(allTasks.map((t, i) => [t.id, i]))

  if (!rows.length) {
    const r = taskStartRow
    setCell(ws, r, 0, allTasks.length ? '(no tasks plotted)' : '(no tasks)', {
      ...styleBodyMuted,
      alignment: { vertical: 'center', horizontal: 'left', indent: 1 },
    })
    for (let c = 1; c < cols; c += 1) {
      setCell(
        ws,
        r,
        c,
        '',
        ganttTaskCellEmpty(false, includeDayColumns && c - 1 === todayIdx),
      )
    }
  } else {
    rows.forEach((t, i) => {
      const r = taskStartRow + i
      const alt = i % 2 === 1
      const depth = depthMap.get(t.id) ?? 0
      const colors = paletteFor(colorIndexById.get(t.id) ?? i)
      const taskRange = ranges.get(t.id) ?? null
      const namePrefix = depth > 0 ? '↳ ' : ''
      const nameSuffix = taskRange ? '' : ' (no dates)'
      setCell(
        ws,
        r,
        0,
        `${namePrefix}${t.name.trim() || '(untitled)'}${nameSuffix}`,
        ganttTaskNameCell(alt, colors.done, depth),
      )

      const taskStartMs = taskRange ? parseISOToUtcMs(taskRange.start) : NaN
      const taskEndMs = taskRange ? parseISOToUtcMs(taskRange.end) : NaN
      const intersects =
        taskRange !== null && taskEndMs >= rangeStartMs && taskStartMs <= rangeEndMs

      let doneDays = 0
      let startOffset = 0
      let endOffset = 0
      if (intersects && taskRange) {
        const visStart = taskStartMs < rangeStartMs ? visibleRange.start : taskRange.start
        const visEnd = taskEndMs > rangeEndMs ? visibleRange.end : taskRange.end
        const spanDays = Math.max(1, daysInclusive(visStart, visEnd))
        const pct = getEffectiveProgress(t, allTasks)
        doneDays = Math.max(0, Math.min(spanDays, Math.round((spanDays * pct) / 100)))
        startOffset = Math.round((parseISOToUtcMs(visStart) - rangeStartMs) / dayMs)
        endOffset = startOffset + spanDays - 1
      }

      if (includeDayColumns) {
        for (let c = 0; c < days.length; c += 1) {
          const colIdx = 1 + c
          const isTodayCol = c === todayIdx
          if (intersects && c >= startOffset && c <= endOffset) {
            const offsetInTask = c - startOffset
            const fillRgb = offsetInTask < doneDays ? colors.done : colors.planned
            setCell(ws, r, colIdx, '', ganttTaskCellFilled(fillRgb))
          } else {
            setCell(ws, r, colIdx, '', ganttTaskCellEmpty(alt, isTodayCol))
          }
        }
      } else {
        let weekDoneRemaining = doneDays
        for (let weekIdx = 0; weekIdx < weekBuckets.length; weekIdx += 1) {
          const week = weekBuckets[weekIdx]!
          const colIdx = 1 + weekIdx
          if (!intersects) {
            setCell(ws, r, colIdx, '', ganttTaskCellEmpty(alt, false))
            continue
          }
          const overlapStart = Math.max(week.startOffset, startOffset)
          const overlapEnd = Math.min(week.endOffset, endOffset)
          const overlapDays = overlapStart <= overlapEnd ? overlapEnd - overlapStart + 1 : 0
          if (overlapDays > 0) {
            const doneInWeek = Math.min(overlapDays, weekDoneRemaining)
            const fillRgb =
              doneInWeek >= overlapDays
                ? colors.done
                : doneInWeek > 0
                  ? colors.partial
                  : colors.planned
            setCell(ws, r, colIdx, '', ganttTaskCellFilled(fillRgb))
            weekDoneRemaining = Math.max(0, weekDoneRemaining - overlapDays)
          } else {
            setCell(ws, r, colIdx, '', ganttTaskCellEmpty(alt, false))
          }
        }
      }
    })
  }

  const lastRow = taskStartRow + Math.max(rows.length, 1) - 1
  applyRange(ws, 0, 0, lastRow, cols - 1)

  const merges: XLSX.Range[] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: cols - 1 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: cols - 1 } },
    { s: { r: taskHeaderTopRow, c: 0 }, e: { r: taskHeaderBottomRow, c: 0 } },
  ]
  if (includeDayColumns) {
    let monthCol = 1
    for (const span of monthSpans) {
      merges.push({
        s: { r: monthHeaderRow, c: monthCol },
        e: { r: monthHeaderRow, c: monthCol + span.dayCount - 1 },
      })
      monthCol += span.dayCount
    }
  } else {
    let weekCol = 1
    while (weekCol <= weekBuckets.length) {
      const firstBucket = weekBuckets[weekCol - 1]!
      const firstDay = days[firstBucket.startOffset] ?? visibleRange.start
      const monthKey = firstDay.slice(0, 7)
      const start = weekCol
      let end = start
      while (end < weekBuckets.length) {
        const nextBucket = weekBuckets[end]!
        const nextFirstDay = days[nextBucket.startOffset] ?? visibleRange.end
        if (nextFirstDay.slice(0, 7) !== monthKey) break
        end += 1
      }
      merges.push({ s: { r: monthHeaderRow, c: start }, e: { r: monthHeaderRow, c: end } })
      weekCol = end + 1
    }
  }
  if (includeDayColumns) {
    let wc = 1
    for (const span of weekSpans) {
      merges.push({
        s: { r: weekHeaderRow, c: wc },
        e: { r: weekHeaderRow, c: wc + span.dayCount - 1 },
      })
      wc += span.dayCount
    }
  }
  ws['!merges'] = merges

  ws['!cols'] = [
    { wch: 32 },
    ...(includeDayColumns
      ? days.map(() => ({ wch: 3 }))
      : weekSpans.map(() => ({ wch: 8 }))),
  ]

  const rowHeights: { hpx: number }[] = []
  rowHeights[0] = { hpx: 32 }
  rowHeights[1] = { hpx: 22 }
  rowHeights[monthHeaderRow] = { hpx: 22 }
  rowHeights[weekHeaderRow] = { hpx: 18 }
  if (includeDayColumns) rowHeights[dayHeaderRow] = { hpx: 16 }
  for (let r = taskStartRow; r <= lastRow; r += 1) {
    rowHeights[r] = { hpx: 24 }
  }
  ws['!rows'] = rowHeights

  return ws
}

/** Writes a workbook (Project info + Tasks + Gantt) and triggers a browser download. */
export function exportGanttToXlsx(
  tasks: GanttTask[],
  options: ExportGanttOptions,
): void {
  const exportedAt = options.exportedAt ?? new Date()
  const exportDate = exportDateYYYYMMDD(exportedAt)
  const displayProjectName =
    options.projectName.trim().length > 0
      ? options.projectName.trim()
      : 'Untitled project'

  const ranges = resolveTaskRanges(tasks)
  // The Gantt sheet mirrors what the app draws: expanded rows that are ticked.
  const plottedRows = getVisibleTasks(tasks).filter(isPlotted)

  const autoRange = unionRanges(plottedRows, ranges)
  const fallbackRange = (() => {
    const start = exportDateYYYYMMDD(exportedAt)
    return { start, end: addDaysISO(start, 30) }
  })()
  const rawRange = options.visibleRange ?? autoRange ?? fallbackRange
  const range = clampRangeOrder(rawRange.start, rawRange.end)

  const projectWs = buildProjectSheet(
    displayProjectName,
    exportDate,
    tasks,
    plottedRows,
    ranges,
    range,
  )
  const tasksWs = buildTasksSheet(displayProjectName, tasks, ranges)
  const ganttWs = buildGanttSheet(
    displayProjectName,
    tasks,
    plottedRows,
    ranges,
    range,
    options.includeDayColumns ?? true,
    exportDate,
    options.weekLabelFormat ?? 'iso',
  )

  const book = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(book, projectWs, 'Project')
  XLSX.utils.book_append_sheet(book, tasksWs, 'Tasks')
  XLSX.utils.book_append_sheet(book, ganttWs, 'Gantt')

  const base = buildGanttExportBasename(options.projectName, exportedAt)
  const filename = `${base}.xlsx`
  XLSX.writeFile(book, filename)
}
