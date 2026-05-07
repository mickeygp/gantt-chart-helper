import {
  type CSSProperties,
  type SetStateAction,
  useEffect,
  useMemo,
  useState,
} from 'react'

import { exportGanttToXlsx } from './exportGanttXlsx'
import { addDaysISO, daysInclusive, parseISOToUtcMs } from './ganttDates'
import { buildMonthSpans, buildWeekSpans } from './ganttTimeline'
import { loadGanttWorkbook, saveGanttWorkbook } from './ganttLocalCache'
import {
  createSheet,
  nextSheetLabel,
  type GanttSheetState,
  type GanttWorkbookState,
} from './ganttSheet'
import { createTask, type GanttTask, todayISO } from './ganttTypes'
import './GanttBuilder.css'

const DAY_PX = 26
const DND_TASK_MIME = 'application/x-gantt-task-id'

const TASK_COLORS = [
  '#6366f1', // indigo
  '#0ea5e9', // sky
  '#10b981', // emerald
  '#f59e0b', // amber
  '#ef4444', // red
  '#ec4899', // pink
  '#8b5cf6', // violet
  '#14b8a6', // teal
]

function taskColor(idx: number): string {
  return TASK_COLORS[idx % TASK_COLORS.length]!
}

function reorderTasks(list: GanttTask[], activeId: string, overId: string): GanttTask[] {
  if (activeId === overId) return list
  const fromIdx = list.findIndex((x) => x.id === activeId)
  const toIdx = list.findIndex((x) => x.id === overId)
  if (fromIdx === -1 || toIdx === -1) return list
  const next = [...list]
  const [moved] = next.splice(fromIdx, 1)
  const insertBefore = next.findIndex((x) => x.id === overId)
  next.splice(insertBefore, 0, moved)
  return next
}

function moveTaskToEnd(list: GanttTask[], activeId: string): GanttTask[] {
  const fromIdx = list.findIndex((x) => x.id === activeId)
  if (fromIdx === -1) return list
  const next = [...list]
  const [moved] = next.splice(fromIdx, 1)
  next.push(moved)
  return next
}

function readDragTaskId(dataTransfer: DataTransfer): string | null {
  const fromMime = dataTransfer.getData(DND_TASK_MIME)
  if (fromMime) return fromMime
  const plain = dataTransfer.getData('text/plain')
  return plain.trim() || null
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

type Props = {
  initialTasks?: GanttTask[]
}

type VisibleRangeShortcut = {
  key: string
  label: string
  title: string
  getRange: () => { start: string; end: string }
}

function clampRangeOrder(start: string, end: string): { start: string; end: string } {
  if (parseISOToUtcMs(end) < parseISOToUtcMs(start)) return { start, end: start }
  return { start, end }
}

function defaultSampleTasks(): GanttTask[] {
  const t = todayISO()
  return [
    createTask({ name: 'Discovery', start: t, end: addDaysISO(t, 6), progress: 100 }),
    createTask({ name: 'Build', start: addDaysISO(t, 7), end: addDaysISO(t, 20), progress: 40 }),
    createTask({ name: 'Launch', start: addDaysISO(t, 21), end: addDaysISO(t, 24), progress: 0 }),
  ]
}

function getInitialWorkbook(initialTasks?: GanttTask[]): GanttWorkbookState {
  if (initialTasks !== undefined) {
    const s = createSheet({ sheetName: 'Sheet 1', tasks: initialTasks })
    return { activeSheetId: s.id, sheets: [s] }
  }
  const wb = loadGanttWorkbook()
  if (wb) return wb
  const first = createSheet({ sheetName: 'Sheet 1', tasks: defaultSampleTasks() })
  return { activeSheetId: first.id, sheets: [first] }
}

export default function GanttBuilder({ initialTasks }: Props) {
  const [workbook, setWorkbook] = useState<GanttWorkbookState>(() =>
    getInitialWorkbook(initialTasks),
  )
  const [renamingSheetId, setRenamingSheetId] = useState<string | null>(null)
  const [renameDraft, setRenameDraft] = useState('')
  const [draggingTaskId, setDraggingTaskId] = useState<string | null>(null)
  const [dragOverTaskId, setDragOverTaskId] = useState<string | null>(null)
  const [dragOverFooter, setDragOverFooter] = useState(false)
  const [includeDayColumnsInExport, setIncludeDayColumnsInExport] = useState(true)

  function endDragSession() {
    setDraggingTaskId(null)
    setDragOverTaskId(null)
    setDragOverFooter(false)
  }

  const { sheets, activeSheetId } = workbook

  const activeSheet = useMemo(
    () => sheets.find((s) => s.id === activeSheetId) ?? sheets[0],
    [sheets, activeSheetId],
  )

  const { projectName, tasks, viewRangeOverride } = activeSheet

  const autoRange = useMemo(() => rangeForTasks(tasks), [tasks])

  const fallbackRange = useMemo(() => {
    const a = todayISO()
    return { start: a, end: addDaysISO(a, 30) }
  }, [])

  const effectiveRange = useMemo(() => {
    if (viewRangeOverride)
      return clampRangeOrder(viewRangeOverride.start, viewRangeOverride.end)
    return autoRange ?? fallbackRange
  }, [viewRangeOverride, autoRange, fallbackRange])

  const timeline = useMemo(() => {
    const days = iterateDays(effectiveRange.start, effectiveRange.end)
    return {
      days,
      totalWidth: days.length * DAY_PX,
      monthSpans: buildMonthSpans(days),
      weekSpans: buildWeekSpans(days),
    }
  }, [effectiveRange])

  const todayOffsetPx = useMemo(() => {
    const today = todayISO()
    if (today < effectiveRange.start || today > effectiveRange.end) return null
    const offsetDays = Math.round(
      (parseISOToUtcMs(today) - parseISOToUtcMs(effectiveRange.start)) / 86_400_000,
    )
    return offsetDays * DAY_PX + DAY_PX / 2
  }, [effectiveRange])

  const visibleRangeShortcuts = useMemo<VisibleRangeShortcut[]>(
    () => [
      {
        key: '14d',
        label: '2 weeks',
        title: 'Set visible timeline to the next 14 days',
        getRange: () => { const start = todayISO(); return { start, end: addDaysISO(start, 13) } },
      },
      {
        key: '30d',
        label: '1 month',
        title: 'Set visible timeline to the next 30 days',
        getRange: () => { const start = todayISO(); return { start, end: addDaysISO(start, 29) } },
      },
      {
        key: '90d',
        label: '1 quarter',
        title: 'Set visible timeline to the next 90 days',
        getRange: () => { const start = todayISO(); return { start, end: addDaysISO(start, 89) } },
      },
      {
        key: '365d',
        label: '1 year',
        title: 'Set visible timeline to the next 365 days',
        getRange: () => { const start = todayISO(); return { start, end: addDaysISO(start, 364) } },
      },
      {
        key: 'this-year',
        label: 'This year',
        title: 'Set visible timeline to this calendar year',
        getRange: () => {
          const year = new Date().getUTCFullYear()
          return { start: `${year}-01-01`, end: `${year}-12-31` }
        },
      },
    ],
    [],
  )

  useEffect(() => {
    if (initialTasks !== undefined) return
    saveGanttWorkbook(workbook)
  }, [initialTasks, workbook])

  function patchActiveSheet(patch: Partial<GanttSheetState>) {
    const id = activeSheetId
    setWorkbook((w) => ({
      ...w,
      sheets: w.sheets.map((s) => (s.id === id ? { ...s, ...patch } : s)),
    }))
  }

  function setTasksState(updater: SetStateAction<GanttTask[]>) {
    const id = activeSheetId
    setWorkbook((w) => ({
      ...w,
      sheets: w.sheets.map((s) => {
        if (s.id !== id) return s
        const next = typeof updater === 'function' ? updater(s.tasks) : updater
        return { ...s, tasks: next }
      }),
    }))
  }

  function updateTask(taskId: string, patch: Partial<GanttTask>) {
    setTasksState((prev) =>
      prev.map((t) => {
        if (t.id !== taskId) return t
        const next = { ...t, ...patch }
        if (parseISOToUtcMs(next.end) < parseISOToUtcMs(next.start)) next.end = next.start
        return next
      }),
    )
  }

  function removeTask(taskId: string) {
    setTasksState((prev) => prev.filter((t) => t.id !== taskId))
  }

  function applyVisibleRangeShortcut(range: { start: string; end: string }) {
    patchActiveSheet({ viewRangeOverride: range })
  }

  function activateSheet(id: string) {
    setRenamingSheetId(null)
    setWorkbook((w) => ({ ...w, activeSheetId: id }))
  }

  function addSheet() {
    setWorkbook((w) => {
      const sheetName = nextSheetLabel(w.sheets)
      const sheet = createSheet({ sheetName, tasks: [] })
      return { activeSheetId: sheet.id, sheets: [...w.sheets, sheet] }
    })
  }

  function removeSheet(id: string) {
    setRenamingSheetId((r) => (r === id ? null : r))
    setWorkbook((w) => {
      if (w.sheets.length <= 1) return w
      const sheets = w.sheets.filter((s) => s.id !== id)
      let nextActive = w.activeSheetId
      if (nextActive === id) {
        const idx = w.sheets.findIndex((s) => s.id === id)
        const neighbor = sheets[Math.max(0, idx - 1)] ?? sheets[0]
        nextActive = neighbor.id
      }
      return { activeSheetId: nextActive, sheets }
    })
  }

  function beginRename(sheet: GanttSheetState) {
    setRenamingSheetId(sheet.id)
    setRenameDraft(sheet.sheetName)
  }

  function commitRename(id: string) {
    const trimmed = (renameDraft.trim() || 'Sheet').slice(0, 31)
    setWorkbook((w) => ({
      ...w,
      sheets: w.sheets.map((s) => (s.id === id ? { ...s, sheetName: trimmed } : s)),
    }))
    setRenamingSheetId(null)
  }

  function cancelRename() {
    setRenamingSheetId(null)
    setRenameDraft('')
  }

  const timelineHeaderStack =
    timeline.days.length === 0 ? null : (
      <div
        className="gantt-chart__header-stack"
        style={{
          width: timeline.totalWidth,
          gridTemplateColumns: `repeat(${timeline.days.length}, ${DAY_PX}px)`,
          gridTemplateRows: 'var(--gantt-header-month-h) var(--gantt-header-week-h)',
        }}
      >
        {(() => {
          let column = 1
          return timeline.monthSpans.map((span, idx) => {
            const start = column
            column += span.dayCount
            return (
              <div
                key={`m-${effectiveRange.start}-${idx}-${span.label}`}
                className="gantt-chart__month-cell"
                style={{ gridColumn: `${start} / span ${span.dayCount}`, gridRow: 1 }}
                title={span.label}
              >
                {span.label}
              </div>
            )
          })
        })()}
        {(() => {
          let column = 1
          return timeline.weekSpans.map((span, idx) => {
            const start = column
            column += span.dayCount
            return (
              <div
                key={`w-${effectiveRange.start}-${idx}-${span.label}`}
                className="gantt-chart__week-cell"
                style={{ gridColumn: `${start} / span ${span.dayCount}`, gridRow: 2 }}
                title={span.label}
              >
                {span.label}
              </div>
            )
          })
        })()}
      </div>
    )

  return (
    <div className="gantt-builder">
      {/* ── App bar ── */}
      <header className="gantt-appbar">
        <div className="gantt-appbar__inner">
          <div className="gantt-appbar__brand">
            <div className="gantt-appbar__icon" aria-hidden="true">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <rect x="3" y="4" width="18" height="4" rx="1" />
                <rect x="3" y="10" width="12" height="4" rx="1" />
                <rect x="3" y="16" width="15" height="4" rx="1" />
              </svg>
            </div>
            <span className="gantt-appbar__wordmark">Gantt</span>
          </div>

          <div className="gantt-appbar__divider" aria-hidden="true" />

          <div className="gantt-appbar__project-wrap">
            <input
              className="gantt-appbar__project-input"
              type="text"
              placeholder="Untitled project"
              value={projectName}
              onChange={(e) => patchActiveSheet({ projectName: e.target.value })}
              maxLength={200}
              aria-label="Project name"
            />
          </div>

          <div className="gantt-appbar__toolbar">
            <select
              className="gantt-input gantt-input-select"
              style={{ fontSize: '13px', padding: '7px 28px 7px 10px' }}
              value={includeDayColumnsInExport ? 'day-week' : 'week-only'}
              onChange={(e) => setIncludeDayColumnsInExport(e.target.value === 'day-week')}
              aria-label="Export timeline detail"
            >
              <option value="day-week">Day + week</option>
              <option value="week-only">Week only</option>
            </select>
            <button
              type="button"
              className="gantt-btn gantt-btn--primary"
              onClick={() =>
                exportGanttToXlsx(tasks, {
                  projectName,
                  visibleRange: effectiveRange,
                  includeDayColumns: includeDayColumnsInExport,
                })
              }
            >
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                <polyline points="7 10 12 15 17 10" />
                <line x1="12" y1="15" x2="12" y2="3" />
              </svg>
              Export XLSX
            </button>
          </div>
        </div>
      </header>

      <div className="gantt-builder__main">
        {/* ── Task table ── */}
        <div className="gantt-section">
          <div className="gantt-section__header">
            <span className="gantt-section__title">Tasks</span>
            {tasks.length > 0 && (
              <span className="gantt-section__count">{tasks.length}</span>
            )}
          </div>

          <section className="gantt-table-wrap" aria-label="Task list">
            <table className="gantt-table">
              <thead>
                <tr>
                  <th scope="col" className="gantt-table__th-drag">
                    <span className="visually-hidden">Reorder</span>
                  </th>
                  <th scope="col">Task</th>
                  <th scope="col">Start</th>
                  <th scope="col">End</th>
                  <th scope="col">Days</th>
                  <th scope="col">Progress</th>
                  <th scope="col">
                    <span className="visually-hidden">Remove</span>
                  </th>
                </tr>
              </thead>
              <tbody>
                {tasks.map((t, idx) => {
                  const color = taskColor(idx)
                  return (
                    <tr
                      key={t.id}
                      className={`gantt-table__task-row${draggingTaskId === t.id ? ' gantt-table__task-row--dragging' : ''}${dragOverTaskId === t.id ? ' gantt-table__task-row--drop-target' : ''}`}
                      style={{ '--task-color': color } as CSSProperties}
                      onDragOver={(e) => {
                        if (!draggingTaskId) return
                        e.preventDefault()
                        e.dataTransfer.dropEffect = 'move'
                        setDragOverTaskId(t.id)
                        setDragOverFooter(false)
                      }}
                      onDrop={(e) => {
                        e.preventDefault()
                        const dragId = readDragTaskId(e.dataTransfer)
                        if (!dragId || dragId === t.id) { endDragSession(); return }
                        setTasksState((prev) => reorderTasks(prev, dragId, t.id))
                        endDragSession()
                      }}
                    >
                      <td className="gantt-table__cell-drag">
                        <button
                          type="button"
                          className="gantt-drag-handle"
                          draggable
                          aria-label={`Drag to reorder row: ${t.name.trim() || 'Untitled task'}`}
                          title="Drag to reorder"
                          onDragStart={(e) => {
                            e.stopPropagation()
                            e.dataTransfer.setData(DND_TASK_MIME, t.id)
                            e.dataTransfer.setData('text/plain', t.id)
                            e.dataTransfer.effectAllowed = 'move'
                            setDraggingTaskId(t.id)
                          }}
                          onDragEnd={endDragSession}
                        >
                          <span className="gantt-drag-handle__glyph" aria-hidden="true">
                            <span className="gantt-drag-handle__dot" />
                            <span className="gantt-drag-handle__dot" />
                            <span className="gantt-drag-handle__dot" />
                            <span className="gantt-drag-handle__dot" />
                            <span className="gantt-drag-handle__dot" />
                            <span className="gantt-drag-handle__dot" />
                          </span>
                        </button>
                      </td>
                      <td>
                        <div className="gantt-task-name-cell">
                          <span className="gantt-task-swatch" aria-hidden="true" />
                          <input
                            className="gantt-input"
                            aria-label={`Name for ${t.name}`}
                            value={t.name}
                            onChange={(e) => updateTask(t.id, { name: e.target.value })}
                          />
                        </div>
                      </td>
                      <td>
                        <input
                          className="gantt-input gantt-input--date"
                          type="date"
                          value={t.start}
                          onChange={(e) => {
                            const start = e.target.value
                            updateTask(t.id, {
                              start,
                              end: parseISOToUtcMs(t.end) < parseISOToUtcMs(start) ? start : t.end,
                            })
                          }}
                        />
                      </td>
                      <td>
                        <input
                          className="gantt-input gantt-input--date"
                          type="date"
                          value={t.end}
                          min={t.start}
                          onChange={(e) => updateTask(t.id, { end: e.target.value })}
                        />
                      </td>
                      <td className="gantt-num">{daysInclusive(t.start, t.end)}</td>
                      <td>
                        <div className="gantt-progress-cell">
                          <input
                            className="gantt-input gantt-input--narrow"
                            type="number"
                            min={0}
                            max={100}
                            aria-label={`Progress for ${t.name}`}
                            value={t.progress}
                            onChange={(e) =>
                              updateTask(t.id, {
                                progress: Math.min(100, Math.max(0, Number(e.target.value) || 0)),
                              })
                            }
                          />
                          <span className="gantt-percent">%</span>
                          <div className="gantt-progress-track" aria-hidden="true">
                            <div
                              className="gantt-progress-track__fill"
                              style={{ width: `${t.progress}%` }}
                            />
                          </div>
                        </div>
                      </td>
                      <td>
                        <button
                          type="button"
                          className="gantt-btn gantt-btn--delete"
                          aria-label={`Remove ${t.name}`}
                          onClick={() => removeTask(t.id)}
                        >
                          ✕
                        </button>
                      </td>
                    </tr>
                  )
                })}
              </tbody>
              <tfoot>
                <tr
                  className={`gantt-table__foot-row${dragOverFooter ? ' gantt-table__foot-row--drop' : ''}`}
                  onDragOver={(e) => {
                    if (!draggingTaskId) return
                    e.preventDefault()
                    e.dataTransfer.dropEffect = 'move'
                    setDragOverFooter(true)
                    setDragOverTaskId(null)
                  }}
                  onDragLeave={(e) => {
                    if (!e.currentTarget.contains(e.relatedTarget as Node))
                      setDragOverFooter(false)
                  }}
                  onDrop={(e) => {
                    e.preventDefault()
                    const dragId = readDragTaskId(e.dataTransfer)
                    if (!dragId) { endDragSession(); return }
                    setTasksState((prev) => moveTaskToEnd(prev, dragId))
                    endDragSession()
                  }}
                >
                  <td colSpan={7}>
                    <div className="gantt-table__foot-inner">
                      <button
                        type="button"
                        className="gantt-btn gantt-btn--secondary gantt-table__add-btn"
                        onClick={() => setTasksState((prev) => [...prev, createTask()])}
                      >
                        + Add task
                      </button>
                      {draggingTaskId ? (
                        <span className="gantt-table__drop-hint" aria-live="polite">
                          Drop here to move to bottom
                        </span>
                      ) : null}
                    </div>
                  </td>
                </tr>
              </tfoot>
            </table>
          </section>
        </div>

        {/* ── Gantt chart ── */}
        <div className="gantt-section">
          <div className="gantt-section__header">
            <span className="gantt-section__title">Timeline</span>
          </div>

          <section className="gantt-chart" aria-label="Gantt timeline">
            <div className="gantt-chart__range">
              <span className="gantt-chart__range-label">View</span>
              <label className="gantt-chart__range-field">
                <span className="visually-hidden">Timeline starts</span>
                <input
                  className="gantt-input gantt-input--date"
                  type="date"
                  value={effectiveRange.start}
                  max={effectiveRange.end}
                  onChange={(e) => {
                    const start = e.target.value
                    patchActiveSheet({
                      viewRangeOverride: clampRangeOrder(start, effectiveRange.end),
                    })
                  }}
                />
              </label>
              <span className="gantt-chart__range-to" aria-hidden="true">–</span>
              <label className="gantt-chart__range-field">
                <span className="visually-hidden">Timeline ends</span>
                <input
                  className="gantt-input gantt-input--date"
                  type="date"
                  value={effectiveRange.end}
                  min={effectiveRange.start}
                  onChange={(e) => {
                    const end = e.target.value
                    patchActiveSheet({
                      viewRangeOverride: clampRangeOrder(effectiveRange.start, end),
                    })
                  }}
                />
              </label>
              <div className="gantt-chart__range-shortcuts" aria-label="Range shortcuts">
                {visibleRangeShortcuts.map((shortcut) => (
                  <button
                    key={shortcut.key}
                    type="button"
                    className="gantt-btn gantt-btn--ghost"
                    title={shortcut.title}
                    onClick={() => applyVisibleRangeShortcut(shortcut.getRange())}
                  >
                    {shortcut.label}
                  </button>
                ))}
              </div>
              <button
                type="button"
                className="gantt-btn gantt-btn--secondary gantt-chart__range-fit"
                title={
                  autoRange
                    ? 'Snap the visible range to your tasks (clears manual dates)'
                    : 'Clear manual range and use the default window until you add tasks'
                }
                onClick={() => patchActiveSheet({ viewRangeOverride: null })}
              >
                Fit all tasks
              </button>
            </div>

            {timeline.days.length === 0 ? (
              <p className="gantt-chart__empty">Set a valid date range.</p>
            ) : tasks.length === 0 ? (
              <>
                <div className="gantt-chart__scroll">
                  <div
                    className="gantt-chart__pan"
                    style={{ '--gantt-day-px': `${DAY_PX}px` } as CSSProperties}
                  >
                    <div className="gantt-chart__label-col">
                      <div className="gantt-chart__label-header-spacer" aria-hidden="true" />
                    </div>
                    <div
                      className="gantt-chart__timeline-col"
                      style={{ width: timeline.totalWidth }}
                    >
                      {timelineHeaderStack}
                      {todayOffsetPx !== null && (
                        <div
                          className="gantt-chart__today-marker"
                          style={{ left: todayOffsetPx }}
                          aria-label="Today"
                        />
                      )}
                    </div>
                  </div>
                </div>
                <p className="gantt-chart__empty gantt-chart__empty--inline">
                  Add a task to see bars on the timeline.
                </p>
              </>
            ) : (
              <div className="gantt-chart__scroll">
                <div
                  className="gantt-chart__pan"
                  style={{ '--gantt-day-px': `${DAY_PX}px` } as CSSProperties}
                >
                  <div className="gantt-chart__label-col">
                    <div className="gantt-chart__label-header-spacer" aria-hidden="true" />
                    {tasks.map((t, idx) => (
                      <div
                        key={t.id}
                        className="gantt-chart__task-name"
                        style={{ '--task-color': taskColor(idx) } as CSSProperties}
                        title={t.name}
                      >
                        <span className="gantt-chart__task-name-dot" aria-hidden="true" />
                        <span className="gantt-chart__task-name-text">
                          {t.name.trim() || 'Untitled'}
                        </span>
                      </div>
                    ))}
                  </div>
                  <div
                    className="gantt-chart__timeline-col"
                    style={{ width: timeline.totalWidth }}
                  >
                    {timelineHeaderStack}
                    {todayOffsetPx !== null && (
                      <div
                        className="gantt-chart__today-marker"
                        style={{ left: todayOffsetPx }}
                        aria-label="Today"
                      />
                    )}
                    {tasks.map((t, idx) => {
                      const rangeStart = timeline.days[0] ?? effectiveRange.start
                      const rangeEnd = timeline.days[timeline.days.length - 1] ?? effectiveRange.end
                      const dayMs = 86_400_000
                      const spanMs = parseISOToUtcMs(rangeEnd) - parseISOToUtcMs(rangeStart)
                      const totalDays = spanMs >= 0 ? Math.floor(spanMs / dayMs) + 1 : 1

                      const r0 = parseISOToUtcMs(rangeStart)
                      const r1 = parseISOToUtcMs(rangeEnd)
                      const t0 = parseISOToUtcMs(t.start)
                      const t1 = parseISOToUtcMs(t.end)
                      const intersects = t1 >= r0 && t0 <= r1

                      let leftPx = 0
                      let widthPx = 0
                      if (intersects) {
                        const visStart = t0 < r0 ? rangeStart : t.start
                        const visEnd = t1 > r1 ? rangeEnd : t.end
                        const offsetDays = Math.round(
                          (parseISOToUtcMs(visStart) - r0) / dayMs,
                        )
                        leftPx = offsetDays * DAY_PX
                        widthPx = Math.max(daysInclusive(visStart, visEnd) * DAY_PX, 4)
                      }

                      return (
                        <div
                          key={t.id}
                          className="gantt-chart__track"
                          style={
                            {
                              width: totalDays * DAY_PX,
                              '--task-color': taskColor(idx),
                            } as CSSProperties
                          }
                        >
                          {intersects ? (
                            <div
                              className="gantt-chart__bar"
                              style={{ left: leftPx, width: widthPx }}
                            >
                              <span
                                className="gantt-chart__bar-fill"
                                style={{ width: `${t.progress}%` }}
                              />
                            </div>
                          ) : null}
                        </div>
                      )
                    })}
                  </div>
                </div>
              </div>
            )}
          </section>
        </div>
      </div>

      {/* ── Sheet tabs ── */}
      <nav className="gantt-tabs" aria-label="Project sheets">
        <div className="gantt-tabs__scroll">
          <div className="gantt-tabs__list" role="tablist" aria-orientation="horizontal">
            {sheets.map((s) => {
              const isActive = s.id === activeSheetId
              const isRenaming = s.id === renamingSheetId
              return (
                <div
                  key={s.id}
                  className={`gantt-tabs__cell${isActive ? ' gantt-tabs__cell--active' : ''}`}
                >
                  {isRenaming ? (
                    <input
                      className="gantt-tabs__rename-input"
                      value={renameDraft}
                      autoFocus
                      maxLength={31}
                      aria-label="Sheet name"
                      onChange={(e) => setRenameDraft(e.target.value)}
                      onBlur={() => commitRename(s.id)}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter') { e.preventDefault(); commitRename(s.id) }
                        if (e.key === 'Escape') { e.preventDefault(); cancelRename() }
                      }}
                    />
                  ) : (
                    <button
                      type="button"
                      role="tab"
                      aria-selected={isActive}
                      id={`gantt-tab-${s.id}`}
                      className="gantt-tabs__tab"
                      title="Double-click to rename"
                      onClick={() => activateSheet(s.id)}
                      onDoubleClick={(e) => { e.preventDefault(); beginRename(s) }}
                    >
                      <span className="gantt-tabs__tab-label">{s.sheetName}</span>
                    </button>
                  )}
                  {sheets.length > 1 ? (
                    <button
                      type="button"
                      className="gantt-tabs__close"
                      aria-label={`Close ${s.sheetName}`}
                      tabIndex={isActive ? 0 : -1}
                      onClick={(e) => { e.stopPropagation(); removeSheet(s.id) }}
                    >
                      ×
                    </button>
                  ) : null}
                </div>
              )
            })}
          </div>
          <button
            type="button"
            className="gantt-tabs__new"
            aria-label="New sheet"
            title="New sheet"
            onClick={addSheet}
          >
            +
          </button>
        </div>
      </nav>
    </div>
  )
}
