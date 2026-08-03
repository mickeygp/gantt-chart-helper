import type { GanttWorkbookState } from './ganttSheet'
import {
  migrateLegacyV1,
  parseWorkbookPayload,
  serializeWorkbook,
} from './ganttSerialize'

const STORAGE_KEY = 'gantt-chart-helper'
/** Legacy payload shape (single project). */
const LEGACY_STORAGE_KEY = 'gantt-chart-helper:v1'

/**
 * localStorage stays the source of truth for signed-out use, and doubles as an
 * offline cache once cloud sync is on — the app always renders from local data
 * first, then reconciles with the server.
 */
export function loadGanttWorkbook(): GanttWorkbookState | null {
  if (typeof localStorage === 'undefined') return null
  try {
    const raw = localStorage.getItem(STORAGE_KEY)
    if (raw) return parseWorkbookPayload(JSON.parse(raw) as unknown)

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
    localStorage.setItem(STORAGE_KEY, JSON.stringify(serializeWorkbook(state)))
  } catch {
    // quota / private mode
  }
}
