/** Parse YYYY-MM-DD as UTC noon to avoid DST edge cases when diffing days. */
export function parseISOToUtcMs(iso: string): number {
  const [y, m, d] = iso.split('-').map(Number)
  if (!y || !m || !d) return NaN
  return Date.UTC(y, m - 1, d, 12, 0, 0)
}

export function daysInclusive(startISO: string, endISO: string): number {
  const a = parseISOToUtcMs(startISO)
  const b = parseISOToUtcMs(endISO)
  if (Number.isNaN(a) || Number.isNaN(b) || b < a) return 0
  return Math.floor((b - a) / 86_400_000) + 1
}

/** Signed whole-day difference, `to` minus `from`. Returns 0 on bad input. */
export function diffDays(fromISO: string, toISO: string): number {
  const a = parseISOToUtcMs(fromISO)
  const b = parseISOToUtcMs(toISO)
  if (Number.isNaN(a) || Number.isNaN(b)) return 0
  return Math.round((b - a) / 86_400_000)
}

export function addDaysISO(iso: string, deltaDays: number): string {
  const t = parseISOToUtcMs(iso)
  if (Number.isNaN(t)) return iso
  const next = new Date(t + deltaDays * 86_400_000)
  const y = next.getUTCFullYear()
  const m = String(next.getUTCMonth() + 1).padStart(2, '0')
  const d = String(next.getUTCDate()).padStart(2, '0')
  return `${y}-${m}-${d}`
}
