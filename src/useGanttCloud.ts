import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import {
  fetchCloudUpdatedAt,
  fetchCloudWorkbook,
  pushCloudWorkbook,
} from './ganttCloud'
import type { GanttWorkbookState } from './ganttSheet'
import { serializeWorkbook } from './ganttSerialize'
import { isCloudConfigured, supabase } from './supabaseClient'

export type CloudStatus =
  | 'off' // no Supabase credentials in this build
  | 'signed-out'
  | 'syncing' // first pull after sign-in
  | 'conflict' // both sides have data and they differ
  | 'saving'
  | 'saved'
  | 'error'

export type CloudSync = {
  configured: boolean
  email: string | null
  status: CloudStatus
  error: string | null
  /** Redirects to Google; the session lands when the browser comes back. */
  signIn: () => Promise<void>
  signOut: () => Promise<void>
  /** Only meaningful while status is 'conflict'. */
  resolveConflict: (keep: 'cloud' | 'local') => void
  dismissError: () => void
}

const PUSH_DEBOUNCE_MS = 1200

function fingerprint(workbook: GanttWorkbookState): string {
  return JSON.stringify(serializeWorkbook(workbook))
}

/**
 * Keeps the workbook mirrored to Supabase for whoever is signed in.
 *
 * Rules of the road:
 *  - localStorage stays authoritative for rendering, so the app never blocks on
 *    the network and still works offline.
 *  - The first pull after sign-in never silently overwrites. If this browser
 *    and the cloud both hold data and they differ, sync pauses and the user
 *    picks a side — losing a plan to a background merge would be worse than
 *    any amount of friction here.
 *  - After that, local edits are debounced upstream, and a window focus checks
 *    whether another device moved ahead.
 *
 * Status is derived rather than stored: `syncedUserId` and `lastPushed` are the
 * only sync state, and every label falls out of comparing them with the current
 * user and workbook. That keeps "is it saved?" honest by construction instead
 * of depending on every code path remembering to update a status field.
 */
export function useGanttCloud(
  workbook: GanttWorkbookState,
  setWorkbook: (next: GanttWorkbookState) => void,
  /** Skips all sync — used when the component is driven by explicit props. */
  disabled: boolean,
): CloudSync {
  const [email, setEmail] = useState<string | null>(null)
  const [userId, setUserId] = useState<string | null>(null)
  const [error, setError] = useState<string | null>(null)

  /** The account whose first pull has completed. Null until it settles. */
  const [syncedUserId, setSyncedUserId] = useState<string | null>(null)
  /** Fingerprint of the payload the server is known to hold. */
  const [lastPushed, setLastPushed] = useState<string | null>(null)
  /** Server timestamp of that payload, used to spot other devices' edits. */
  const [lastUpdatedAt, setLastUpdatedAt] = useState<string | null>(null)
  /**
   * The cloud side of an unresolved conflict, tagged with the account it came
   * from, so a sign-out or account switch invalidates it on its own.
   */
  const [pendingCloud, setPendingCloud] = useState<{
    userId: string
    workbook: GanttWorkbookState
  } | null>(null)

  const workbookRef = useRef(workbook)
  useEffect(() => {
    workbookRef.current = workbook
  }, [workbook])

  const active = isCloudConfigured && !disabled
  const print = useMemo(() => fingerprint(workbook), [workbook])

  const armed = userId !== null && syncedUserId === userId
  const conflicted = userId !== null && pendingCloud?.userId === userId
  const dirty = armed && lastPushed !== null && print !== lastPushed

  const status: CloudStatus = !active
    ? 'off'
    : !userId
      ? 'signed-out'
      : error
        ? 'error'
        : conflicted
          ? 'conflict'
          : !armed
            ? 'syncing'
            : dirty
              ? 'saving'
              : 'saved'

  // ── Session tracking ──────────────────────────────────────────────────────

  useEffect(() => {
    if (!active || !supabase) return
    let cancelled = false

    void supabase.auth.getSession().then(({ data }) => {
      if (cancelled) return
      setUserId(data.session?.user.id ?? null)
      setEmail(data.session?.user.email ?? null)
    })

    const { data: sub } = supabase.auth.onAuthStateChange((_event, session) => {
      setUserId(session?.user.id ?? null)
      setEmail(session?.user.email ?? null)
    })

    return () => {
      cancelled = true
      sub.subscription.unsubscribe()
    }
  }, [active])

  // ── First pull after sign-in ──────────────────────────────────────────────

  useEffect(() => {
    if (!active || !userId || syncedUserId === userId) return
    let cancelled = false

    void fetchCloudWorkbook(userId).then((res) => {
      if (cancelled) return
      if (!res.ok) {
        setError(res.error)
        return
      }

      const local = workbookRef.current
      // Nothing saved yet: this browser seeds the account.
      if (!res.data) {
        void pushCloudWorkbook(userId, local).then((push) => {
          if (cancelled) return
          if (!push.ok) {
            setError(push.error)
            return
          }
          setLastPushed(fingerprint(local))
          setLastUpdatedAt(push.data)
          setSyncedUserId(userId)
        })
        return
      }

      const cloudPrint = fingerprint(res.data.workbook)
      setLastUpdatedAt(res.data.updatedAt)
      if (cloudPrint === fingerprint(local)) {
        setLastPushed(cloudPrint)
        setSyncedUserId(userId)
        return
      }

      // Two different plans. Ask rather than guess.
      setPendingCloud({ userId, workbook: res.data.workbook })
    })

    return () => {
      cancelled = true
    }
  }, [active, userId, syncedUserId])

  const resolveConflict = useCallback(
    (keep: 'cloud' | 'local') => {
      if (!userId || pendingCloud?.userId !== userId) return
      const cloud = pendingCloud.workbook
      setPendingCloud(null)

      if (keep === 'cloud') {
        setLastPushed(fingerprint(cloud))
        setSyncedUserId(userId)
        setWorkbook(cloud)
        return
      }

      const local = workbookRef.current
      void pushCloudWorkbook(userId, local).then((push) => {
        if (!push.ok) {
          setError(push.error)
          return
        }
        setLastPushed(fingerprint(local))
        setLastUpdatedAt(push.data)
        setSyncedUserId(userId)
      })
    },
    [pendingCloud, setWorkbook, userId],
  )

  // ── Debounced push of local edits ─────────────────────────────────────────

  useEffect(() => {
    if (!active || !userId || !armed || !dirty) return
    const timer = setTimeout(() => {
      void pushCloudWorkbook(userId, workbookRef.current).then((push) => {
        if (!push.ok) {
          setError(push.error)
          return
        }
        setLastPushed(fingerprint(workbookRef.current))
        setLastUpdatedAt(push.data)
      })
    }, PUSH_DEBOUNCE_MS)

    return () => clearTimeout(timer)
  }, [active, userId, armed, dirty, print])

  // ── Notice edits made on another device ───────────────────────────────────

  useEffect(() => {
    if (!active || !userId || !armed) return

    function onFocus() {
      if (!userId) return
      // A local edit is still unsaved; let the push win rather than racing it.
      if (fingerprint(workbookRef.current) !== lastPushed) return

      void fetchCloudUpdatedAt(userId).then((res) => {
        if (!res.ok || !res.data || res.data === lastUpdatedAt) return
        void fetchCloudWorkbook(userId).then((full) => {
          if (!full.ok || !full.data) return
          // Re-check: the user may have started editing while we were fetching.
          if (fingerprint(workbookRef.current) !== lastPushed) return
          setLastPushed(fingerprint(full.data.workbook))
          setLastUpdatedAt(full.data.updatedAt)
          setWorkbook(full.data.workbook)
        })
      })
    }

    window.addEventListener('focus', onFocus)
    return () => window.removeEventListener('focus', onFocus)
  }, [active, armed, lastPushed, lastUpdatedAt, setWorkbook, userId])

  // ── Auth actions ──────────────────────────────────────────────────────────

  /**
   * Google OAuth rather than an emailed link: no mail is sent, so there is no
   * SMTP to configure and no spam filter between a teammate and their account.
   * This call navigates away — the session is picked up from the callback URL
   * when the browser returns (`detectSessionInUrl` in supabaseClient.ts).
   */
  const signIn = useCallback(async () => {
    if (!supabase) return
    setError(null)
    const { error: err } = await supabase.auth.signInWithOAuth({
      provider: 'google',
      options: { redirectTo: window.location.origin + window.location.pathname },
    })
    if (err) setError(err.message)
  }, [])

  const signOut = useCallback(async () => {
    if (!supabase) return
    await supabase.auth.signOut()
    setSyncedUserId(null)
    setLastPushed(null)
    setLastUpdatedAt(null)
    setPendingCloud(null)
  }, [])

  const dismissError = useCallback(() => setError(null), [])

  return {
    configured: isCloudConfigured,
    email,
    status,
    error,
    signIn,
    signOut,
    resolveConflict,
    dismissError,
  }
}
