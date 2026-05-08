import { addDaysISO, parseISOToUtcMs } from './ganttDates'

export type TimelineSpan = {
  label: string
  dayCount: number
}

/** Monday (UTC) of the calendar week containing `iso` (YYYY-MM-DD). */
export function mondayOfWeek(iso: string): string {
  const ms = parseISOToUtcMs(iso)
  const dow = new Date(ms).getUTCDay()
  const toMonday = dow === 0 ? -6 : 1 - dow
  return addDaysISO(iso, toMonday)
}

export function formatMonthYear(isoFirstDay: string): string {
  const t = parseISOToUtcMs(isoFirstDay)
  if (Number.isNaN(t)) return isoFirstDay
  return new Intl.DateTimeFormat(undefined, {
    month: 'short',
    year: 'numeric',
    timeZone: 'UTC',
  }).format(new Date(t))
}

/** ISO 8601 week number for the Monday of a week (pass the Monday ISO date). */
export function isoWeekNumber(mondayISO: string): number {
  const ms = parseISOToUtcMs(mondayISO)
  const d = new Date(ms)
  // Thursday of this week — ISO weeks are defined by their Thursday
  const thu = new Date(ms + (4 - (d.getUTCDay() || 7)) * 86_400_000)
  const yearStart = Date.UTC(thu.getUTCFullYear(), 0, 1)
  return Math.ceil(((thu.getTime() - yearStart) / 86_400_000 + 1) / 7)
}

/** "Week N" label for a week band, derived from the Monday of that week. */
export function formatWeekLabel(mondayISO: string): string {
  return `Week ${isoWeekNumber(mondayISO)}`
}

export function buildMonthSpans(days: string[]): TimelineSpan[] {
  if (!days.length) return []
  const spans: TimelineSpan[] = []
  let i = 0
  while (i < days.length) {
    const ysm = days[i].slice(0, 7)
    let j = i + 1
    while (j < days.length && days[j].slice(0, 7) === ysm) j++
    spans.push({
      label: formatMonthYear(days[i]),
      dayCount: j - i,
    })
    i = j
  }
  return spans
}

export function buildWeekSpans(days: string[]): TimelineSpan[] {
  if (!days.length) return []
  const spans: TimelineSpan[] = []
  let i = 0
  while (i < days.length) {
    const wkMon = mondayOfWeek(days[i])
    let j = i + 1
    while (j < days.length && mondayOfWeek(days[j]) === wkMon) j++
    spans.push({
      label: formatWeekLabel(wkMon),
      dayCount: j - i,
    })
    i = j
  }
  return spans
}
