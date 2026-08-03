# Gantt Chart Helper

A browser-based Gantt planner: task table on top, draggable timeline below,
multiple sheets per workbook, and XLSX export.

```bash
npm install
npm run dev
```

## Sheets

Each tab is an independent plan — its own tasks, project name, and visible date
range. Use the ⧉ button on a tab (or <kbd>Shift</kbd>+<kbd>D</kbd>) to duplicate
it. The copy is fully detached: task ids are regenerated, and parent links and
dependencies are rewritten to point inside the copy, so editing "Plan B" never
touches "Plan A".

## Linked tasks

The **Depends on** column links a task to the ones it runs after
(finish-to-start). Moving a predecessor moves everything downstream by the same
number of days, **preserving the gap** you left between them — a two-day buffer
stays two days rather than collapsing.

What triggers a ripple:

| Action | Successors move? |
| --- | --- |
| Drag a bar sideways | Yes |
| Drag the bar's **right** edge | Yes — the finish date moved |
| Drag the bar's **left** edge | No — duration changed, finish did not |
| Type a new **End** date | Yes |
| Type a new **Start** date | No |

Parent/child rollup is respected in both directions: dragging a subtask moves
its parent's summary bar, so anything depending on the parent ripples too, and
shifting a linked parent carries its whole subtree along.

The picker only offers tasks that cannot create a loop — ancestors, descendants,
and anything already downstream are filtered out, so there is no invalid state
to recover from. Links are drawn on the timeline as connectors between bars.

## Cloud sync (optional)

Without credentials the app is localStorage-only and the cloud UI stays hidden.
With Supabase configured you get sign-in, cross-device sync, and share links.

**1. Create a Supabase project** at [supabase.com](https://supabase.com).

**2. Run the schema.** Dashboard → SQL Editor → New query → paste
[`supabase/schema.sql`](supabase/schema.sql) → Run. It is idempotent.

**3. Add credentials.** Copy `.env.example` to `.env.local` and fill in the
project URL and anon key from Project Settings → API Keys. For the Vercel
deploy, add the same two variables in Project Settings → Environment Variables
and redeploy.

**4. Enable Google sign-in.** Dashboard → Authentication → Providers → Google.
Paste in a client id and secret from a Google Cloud OAuth 2.0 credential, and
add the callback URL Supabase shows you to that credential's *Authorized
redirect URIs*.

**5. Set the redirect URLs.** Dashboard → Authentication → URL Configuration →
Redirect URLs. Add `http://localhost:5173/**` and your deployed origin, or
sign-in will bounce on the way back. Note the `/**` — a bare origin misses the
trailing slash.

Sign-in is Google OAuth, so **no email is ever sent** and no SMTP is needed. If
you would rather use emailed magic links, swap `signInWithOAuth` for
`signInWithOtp` in [`src/useGanttCloud.ts`](src/useGanttCloud.ts) — but budget
time for configuring custom SMTP, because Supabase's built-in mailer is rate
limited and gets filtered by corporate inboxes.

### How syncing behaves

- localStorage stays the rendering source of truth, so the app never blocks on
  the network and still works offline.
- Sign-in is one click through Google — no password is ever typed into the app.
- Edits are pushed on a debounce; the pill in the header reads
  Synced / Saving… / Sync paused.
- Refocusing the tab checks whether another device moved ahead, and pulls if so
  — but never while you have unsaved local edits.
- **The first pull after sign-in never silently overwrites.** If your account
  and this browser both hold a workbook and they differ, sync pauses and asks
  which to keep.

### Share links

**Share** publishes a snapshot of the current sheet and copies a link. Opening
that link imports the sheet as a new tab in the recipient's workbook, re-keyed
so it cannot collide with what they already have.

A snapshot is a point-in-time copy, not a live view: later edits do not change
an already-sent link — re-share to publish them. Anyone holding the link can
read that snapshot, so treat it as unlisted-but-public. Links are served through
a `security definer` function that returns exactly one row for an exact id, so a
link holder cannot enumerate anyone else's shares.

## Keyboard shortcuts

Press <kbd>?</kbd> in the app for the full list.
