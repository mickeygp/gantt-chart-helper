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

/** Week stripe label: day range visible in the timeline slice. */
export function formatWeekBandLabel(spanFirst: string, spanLast: string): string {
  const a = parseISOToUtcMs(spanFirst)
  const b = parseISOToUtcMs(spanLast)
  if (Number.isNaN(a) || Number.isNaN(b)) return ''
  const m0 = spanFirst.slice(5, 7)
  const m1 = spanLast.slice(5, 7)
  const d0 = Number(spanFirst.slice(8, 10))
  const d1 = Number(spanLast.slice(8, 10))
  if (m0 === m1 && spanFirst.slice(0, 7) === spanLast.slice(0, 7)) {
    return `${d0}–${d1}`
  }
  const dm = new Intl.DateTimeFormat(undefined, {
    month: 'short',
    day: 'numeric',
    timeZone: 'UTC',
  })
  return `${dm.format(new Date(a))} – ${dm.format(new Date(b))}`
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
      label: formatWeekBandLabel(days[i], days[j - 1]),
      dayCount: j - i,
    })
    i = j
  }
  return spans
}
