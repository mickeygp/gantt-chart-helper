import type { GanttSheetState, GanttWorkbookState } from './ganttSheet'
import { parseSheet, parseWorkbookPayload, serializeWorkbook } from './ganttSerialize'
import { supabase } from './supabaseClient'

/**
 * Reads and writes for the two cloud tables (see supabase/schema.sql):
 *
 *   workbooks      one row per user — the whole workbook as a JSON payload
 *   shared_sheets  a published snapshot of one sheet, readable by link
 *
 * Every response is run back through the same validators the local cache uses.
 * A row is just JSON the server handed us; it gets no more trust than
 * localStorage does.
 */

export type CloudWorkbook = {
  workbook: GanttWorkbookState
  updatedAt: string
}

export type CloudResult<T> =
  | { ok: true; data: T }
  | { ok: false; error: string }

/**
 * Supabase reports query failures as plain `{ message, details, hint, code }`
 * objects rather than Error instances, so `String(e)` would put a literal
 * "[object Object]" in front of the user.
 */
function describe(e: unknown): string {
  if (e instanceof Error) return e.message
  if (typeof e === 'string') return e
  if (e && typeof e === 'object' && 'message' in e) {
    const message = (e as { message: unknown }).message
    if (typeof message === 'string' && message) return message
  }
  return 'unexpected error'
}

function fail(message: string, e: unknown): { ok: false; error: string } {
  return { ok: false, error: `${message}: ${describe(e)}` }
}

/** Fetches the signed-in user's workbook. `data: null` means none saved yet. */
export async function fetchCloudWorkbook(
  userId: string,
): Promise<CloudResult<CloudWorkbook | null>> {
  if (!supabase) return { ok: false, error: 'Cloud sync is not configured' }
  try {
    const { data, error } = await supabase
      .from('workbooks')
      .select('payload, updated_at')
      .eq('user_id', userId)
      .maybeSingle()
    if (error) throw error
    if (!data) return { ok: true, data: null }

    const workbook = parseWorkbookPayload(data.payload)
    if (!workbook) return { ok: false, error: 'Saved workbook could not be read' }
    return { ok: true, data: { workbook, updatedAt: data.updated_at as string } }
  } catch (e) {
    return fail('Could not load from cloud', e)
  }
}

/** Upserts the whole workbook. Returns the server's new updated_at. */
export async function pushCloudWorkbook(
  userId: string,
  workbook: GanttWorkbookState,
): Promise<CloudResult<string>> {
  if (!supabase) return { ok: false, error: 'Cloud sync is not configured' }
  try {
    const { data, error } = await supabase
      .from('workbooks')
      .upsert(
        {
          user_id: userId,
          payload: serializeWorkbook(workbook),
          updated_at: new Date().toISOString(),
        },
        { onConflict: 'user_id' },
      )
      .select('updated_at')
      .single()
    if (error) throw error
    return { ok: true, data: data.updated_at as string }
  } catch (e) {
    return fail('Could not save to cloud', e)
  }
}

/** Cheap poll used on window focus to notice edits made on another device. */
export async function fetchCloudUpdatedAt(
  userId: string,
): Promise<CloudResult<string | null>> {
  if (!supabase) return { ok: false, error: 'Cloud sync is not configured' }
  try {
    const { data, error } = await supabase
      .from('workbooks')
      .select('updated_at')
      .eq('user_id', userId)
      .maybeSingle()
    if (error) throw error
    return { ok: true, data: (data?.updated_at as string | undefined) ?? null }
  } catch (e) {
    return fail('Could not check cloud', e)
  }
}

// ── Share links ─────────────────────────────────────────────────────────────

/**
 * Publishes a snapshot of one sheet and returns its share id. The snapshot is
 * a copy, not a live view: re-share to publish later edits. That keeps a link
 * you sent to someone from changing under them while they read it.
 */
export async function publishSharedSheet(
  sheet: GanttSheetState,
): Promise<CloudResult<string>> {
  if (!supabase) return { ok: false, error: 'Cloud sync is not configured' }
  try {
    const { data, error } = await supabase
      .from('shared_sheets')
      .insert({ sheet_name: sheet.sheetName, sheet })
      .select('id')
      .single()
    if (error) throw error
    return { ok: true, data: data.id as string }
  } catch (e) {
    return fail('Could not create share link', e)
  }
}

/**
 * Loads a shared sheet by id. No auth required — the id is the capability.
 * Goes through the `get_shared_sheet` function rather than selecting the table
 * directly, so a link holder can read their one row without being able to list
 * everyone else's.
 */
export async function fetchSharedSheet(
  shareId: string,
): Promise<CloudResult<GanttSheetState>> {
  if (!supabase) return { ok: false, error: 'Cloud sync is not configured' }
  try {
    const { data, error } = await supabase.rpc('get_shared_sheet', {
      share_id: shareId,
    })
    if (error) throw error
    if (!data) return { ok: false, error: 'That share link no longer exists' }

    const sheet = parseSheet(data)
    if (!sheet) return { ok: false, error: 'That shared sheet could not be read' }
    return { ok: true, data: sheet }
  } catch (e) {
    return fail('Could not open share link', e)
  }
}
