import * as XLSX from 'xlsx-js-style'

import type { GanttTask } from './ganttTypes'
import { addDaysISO, daysInclusive, parseISOToUtcMs } from './ganttDates'
import { buildMonthSpans, buildWeekSpans, formatMonthYear } from './ganttTimeline'

export type ExportGanttOptions = {
  projectName: string
  visibleRange?: { start: string; end: string }
  includeDayColumns?: boolean
  /** Used for filename and workbook metadata; defaults to `new Date()`. */
  exportedAt?: Date
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

function taskRows(tasks: GanttTask[]) {
  return tasks.map((t) => ({
    Task: t.name.trim() || '(untitled)',
    Start: t.start,
    End: t.end,
    'Duration (days)': Math.max(0, daysInclusive(t.start, t.end)),
    'Progress %': Math.min(100, Math.max(0, Math.round(t.progress))),
  }))
}

function clampRangeOrder(start: string, end: string): { start: string; end: string } {
  if (parseISOToUtcMs(end) < parseISOToUtcMs(start)) return { start, end: start }
  return { start, end }
}

function rangeForTasks(tasks: GanttTask[]): { start: string; end: string } | null {
  if (!tasks.length) return null
  let min = tasks[0].start
  let max = tasks[0].end
  for (const t of tasks) {
    if (t.start < min) min = t.start
    if (t.end > max) max = t.end
  }
  return { start: min, end: max }
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

function buildGanttSheet(
  tasks: GanttTask[],
  visibleRange: { start: string; end: string },
  includeDayColumns: boolean,
): XLSX.WorkSheet {
  const days = iterateDays(visibleRange.start, visibleRange.end)
  const monthSpans = buildMonthSpans(days)
  const weekSpans = buildWeekSpans(days)
  const weekBuckets = (() => {
    const buckets: { label: string; dayCount: number; startOffset: number; endOffset: number }[] = []
    let startOffset = 0
    for (const span of weekSpans) {
      const endOffset = startOffset + span.dayCount - 1
      buckets.push({
        label: span.label,
        dayCount: span.dayCount,
        startOffset,
        endOffset,
      })
      startOffset = endOffset + 1
    }
    return buckets
  })()
  const headerColCount = 1 + (includeDayColumns ? days.length : weekSpans.length)
  const aoa: (string | number)[][] = []
  const filledCells: { r: number; c: number; fill: { fgColor: { rgb: string } } }[] = []

  const FILL_DONE = { fgColor: { rgb: 'FF22C55E' } } // green
  const FILL_PARTIAL = { fgColor: { rgb: 'FFFACC15' } } // yellow
  const FILL_PLANNED = { fgColor: { rgb: 'FFD1D5DB' } } // grey

  const monthRow = Array<string | number>(headerColCount).fill('')
  monthRow[0] = 'Task'
  if (includeDayColumns) {
    let monthStartCol = 1
    for (const span of monthSpans) {
      monthRow[monthStartCol] = span.label
      monthStartCol += span.dayCount
    }
  } else {
    let monthStartCol = 1
    while (monthStartCol <= weekBuckets.length) {
      const firstBucket = weekBuckets[monthStartCol - 1]
      const firstDay = days[firstBucket.startOffset] ?? visibleRange.start
      const monthKey = firstDay.slice(0, 7)
      monthRow[monthStartCol] = formatMonthYear(`${monthKey}-01`)
      let monthEndCol = monthStartCol
      while (monthEndCol < weekBuckets.length) {
        const nextBucket = weekBuckets[monthEndCol]
        const nextFirstDay = days[nextBucket.startOffset] ?? visibleRange.end
        if (nextFirstDay.slice(0, 7) !== monthKey) break
        monthEndCol += 1
      }
      monthStartCol = monthEndCol + 1
    }
  }
  aoa.push(monthRow)

  const weekRow = Array<string | number>(headerColCount).fill('')
  weekRow[0] = ''
  let weekStartCol = 1
  for (const bucket of weekBuckets) {
    weekRow[weekStartCol] = bucket.label
    weekStartCol += includeDayColumns ? bucket.dayCount : 1
  }
  aoa.push(weekRow)

  if (includeDayColumns) {
    const dayRow = Array<string | number>(headerColCount).fill('')
    dayRow[0] = ''
    days.forEach((d, idx) => {
      dayRow[idx + 1] = Number(d.slice(8, 10))
    })
    aoa.push(dayRow)
  }

  const rangeStartMs = parseISOToUtcMs(visibleRange.start)
  const rangeEndMs = parseISOToUtcMs(visibleRange.end)
  const dayMs = 86_400_000

  for (const t of tasks) {
    const row = Array<string | number>(headerColCount).fill('')
    row[0] = t.name.trim() || '(untitled)'
    const rowIndex = aoa.length

    const taskStartMs = parseISOToUtcMs(t.start)
    const taskEndMs = parseISOToUtcMs(t.end)
    const intersects = taskEndMs >= rangeStartMs && taskStartMs <= rangeEndMs

    if (intersects) {
      const visStart = taskStartMs < rangeStartMs ? visibleRange.start : t.start
      const visEnd = taskEndMs > rangeEndMs ? visibleRange.end : t.end
      const spanDays = Math.max(1, daysInclusive(visStart, visEnd))
      const doneDays = Math.max(0, Math.min(spanDays, Math.round((spanDays * t.progress) / 100)))
      if (includeDayColumns) {
        const startOffset = Math.round((parseISOToUtcMs(visStart) - rangeStartMs) / dayMs)
        for (let i = 0; i < spanDays; i += 1) {
          const colIndex = 1 + startOffset + i
          row[colIndex] = ''
          filledCells.push({
            r: rowIndex,
            c: colIndex,
            fill: i < doneDays ? FILL_DONE : FILL_PLANNED,
          })
        }
      } else {
        let weekDoneRemaining = doneDays
        for (let weekIdx = 0; weekIdx < weekBuckets.length; weekIdx += 1) {
          const week = weekBuckets[weekIdx]
          const taskStartOffset = Math.round((parseISOToUtcMs(visStart) - rangeStartMs) / dayMs)
          const taskEndOffset = taskStartOffset + spanDays - 1
          const overlapStart = Math.max(week.startOffset, taskStartOffset)
          const overlapEnd = Math.min(week.endOffset, taskEndOffset)
          const overlapDays =
            overlapStart <= overlapEnd ? overlapEnd - overlapStart + 1 : 0
          if (overlapDays > 0) {
            const doneInWeek = Math.min(overlapDays, weekDoneRemaining)
            const colIndex = 1 + weekIdx
            row[colIndex] = ''
            filledCells.push({
              r: rowIndex,
              c: colIndex,
              fill:
                doneInWeek >= overlapDays
                  ? FILL_DONE
                  : doneInWeek > 0
                    ? FILL_PARTIAL
                    : FILL_PLANNED,
            })
            weekDoneRemaining = Math.max(0, weekDoneRemaining - overlapDays)
          }
        }
      }
    }

    aoa.push(row)
  }

  if (!tasks.length) {
    const row = Array<string | number>(headerColCount).fill('')
    row[0] = 'No tasks'
    aoa.push(row)
  }

  const ws = XLSX.utils.aoa_to_sheet(aoa)
  const merges: XLSX.Range[] = []
  let monthCol = 1
  if (includeDayColumns) {
    for (const span of monthSpans) {
      merges.push({
        s: { r: 0, c: monthCol },
        e: { r: 0, c: monthCol + span.dayCount - 1 },
      })
      monthCol += span.dayCount
    }
  } else {
    let weekCol = 1
    while (weekCol <= weekBuckets.length) {
      const firstBucket = weekBuckets[weekCol - 1]
      const firstDay = days[firstBucket.startOffset] ?? visibleRange.start
      const monthKey = firstDay.slice(0, 7)
      const start = weekCol
      let end = start
      while (end < weekBuckets.length) {
        const nextBucket = weekBuckets[end]
        const nextFirstDay = days[nextBucket.startOffset] ?? visibleRange.end
        if (nextFirstDay.slice(0, 7) !== monthKey) break
        end += 1
      }
      merges.push({ s: { r: 0, c: start }, e: { r: 0, c: end } })
      weekCol = end + 1
    }
  }
  let weekCol = 1
  for (const span of weekSpans) {
    merges.push({
      s: { r: 1, c: weekCol },
      e: { r: 1, c: weekCol + (includeDayColumns ? span.dayCount - 1 : 0) },
    })
    weekCol += includeDayColumns ? span.dayCount : 1
  }
  ws['!merges'] = merges
  ws['!cols'] = [
    { wch: 28 },
    ...(includeDayColumns
      ? days.map(() => ({ wch: 2.5 }))
      : weekSpans.map(() => ({ wch: 7.5 }))),
  ]

  // Use cell background fills instead of block glyphs for Gantt bars.
  for (const cell of filledCells) {
    const addr = XLSX.utils.encode_cell({ r: cell.r, c: cell.c })
    const target = ws[addr]
    if (!target) continue
    target.s = {
      ...(target.s ?? {}),
      fill: {
        patternType: 'solid',
        ...cell.fill,
      },
    }
  }

  return ws
}

/** Writes a workbook (Project info + Tasks) and triggers a browser download. */
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

  const projectSheetData: (string | number)[][] = [
    ['Field', 'Value'],
    ['Project name', displayProjectName],
    ['Export date', exportDate],
    ['Tasks exported', tasks.length],
  ]
  const projectWs = XLSX.utils.aoa_to_sheet(projectSheetData)
  projectWs['!cols'] = [{ wch: 18 }, { wch: 44 }]

  const rows = taskRows(tasks)
  const tasksWs = XLSX.utils.json_to_sheet(
    rows.length
      ? rows
      : [
          {
            Task: '',
            Start: '',
            End: '',
            'Duration (days)': '',
            'Progress %': '',
          },
        ],
  )
  tasksWs['!cols'] = [{ wch: 36 }, { wch: 14 }, { wch: 14 }, { wch: 18 }, { wch: 12 }]

  const autoRange = rangeForTasks(tasks)
  const fallbackRange = (() => {
    const now = new Date()
    const y = now.getFullYear()
    const m = String(now.getMonth() + 1).padStart(2, '0')
    const d = String(now.getDate()).padStart(2, '0')
    const start = `${y}-${m}-${d}`
    return { start, end: addDaysISO(start, 30) }
  })()
  const rawRange = options.visibleRange ?? autoRange ?? fallbackRange
  const range = clampRangeOrder(rawRange.start, rawRange.end)
  const ganttWs = buildGanttSheet(tasks, range, options.includeDayColumns ?? true)

  const book = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(book, projectWs, 'Project')
  XLSX.utils.book_append_sheet(book, tasksWs, 'Tasks')
  XLSX.utils.book_append_sheet(book, ganttWs, 'Gantt')

  const base = buildGanttExportBasename(options.projectName, exportedAt)
  const filename = `${base}.xlsx`
  XLSX.writeFile(book, filename)
}
