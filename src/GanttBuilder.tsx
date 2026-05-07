import {
  useCallback,
  type CSSProperties,
  type PointerEvent as ReactPointerEvent,
  type SetStateAction,
  useEffect,
  useMemo,
  useRef,
  useState,
} from 'react'

type Theme = 'light' | 'dark' | 'system'

function useTheme(): [Theme, () => void] {
  const [theme, setTheme] = useState<Theme>(() => {
    const stored = localStorage.getItem('gantt-theme') as Theme | null
    return stored ?? 'system'
  })

  useEffect(() => {
    const root = document.documentElement
    if (theme === 'dark') {
      root.classList.add('dark')
      root.classList.remove('light')
    } else if (theme === 'light') {
      root.classList.add('light')
      root.classList.remove('dark')
    } else {
      root.classList.remove('dark', 'light')
    }
    if (theme === 'system') {
      localStorage.removeItem('gantt-theme')
    } else {
      localStorage.setItem('gantt-theme', theme)
    }
  }, [theme])

  const cycle = useCallback(() => {
    setTheme((t) => (t === 'system' ? 'dark' : t === 'dark' ? 'light' : 'system'))
  }, [])

  return [theme, cycle]
}

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
type BarDragMode = 'move' | 'resize-start' | 'resize-end'

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
  const activeTask = list.find((x) => x.id === activeId)
  if (!activeTask) return list
  const fromIdx = list.findIndex((x) => x.id === activeId)
  const toIdx = list.findIndex((x) => x.id === overId)
  if (fromIdx === -1 || toIdx === -1) return list

  // Subtasks must stay within their parent's subtree
  if (activeTask.parentId) {
    const parentIdx = list.findIndex((x) => x.id === activeTask.parentId)
    if (parentIdx === -1) return list
    const lastDescIdx = findLastDescendantIndex(list, activeTask.parentId)
    if (toIdx <= parentIdx || toIdx > lastDescIdx) return list
  }

  // Collect the block: active task + all its descendants (keeps children together)
  const descendantIds = collectDescendantIds(list, activeId)
  const blockIds = new Set([activeId, ...descendantIds])
  if (blockIds.has(overId)) return list

  const block = list.filter((t) => blockIds.has(t.id))
  const remaining = list.filter((t) => !blockIds.has(t.id))
  const insertAt = remaining.findIndex((x) => x.id === overId)
  if (insertAt === -1) return [...remaining, ...block]
  remaining.splice(insertAt, 0, ...block)
  return remaining
}

function moveTaskToEnd(list: GanttTask[], activeId: string): GanttTask[] {
  const activeTask = list.find((x) => x.id === activeId)
  if (!activeTask) return list
  // Subtasks must not leave their parent's group
  if (activeTask.parentId) return list
  const descendantIds = collectDescendantIds(list, activeId)
  const blockIds = new Set([activeId, ...descendantIds])
  const block = list.filter((t) => blockIds.has(t.id))
  const remaining = list.filter((t) => !blockIds.has(t.id))
  return [...remaining, ...block]
}

function countAncestorDepth(tasks: GanttTask[], task: GanttTask): number {
  const byId = new Map(tasks.map((x) => [x.id, x]))
  let depth = 0
  let curParentId = task.parentId
  while (curParentId) {
    const parent = byId.get(curParentId)
    if (!parent) break
    depth += 1
    if (depth > 12) break
    curParentId = parent.parentId
  }
  return depth
}

function collectDescendantIds(tasks: GanttTask[], parentId: string): Set<string> {
  const descendantIds = new Set<string>()
  const queue = [parentId]
  while (queue.length > 0) {
    const currentId = queue.shift()
    if (!currentId) continue
    for (const t of tasks) {
      if (t.parentId !== currentId) continue
      if (descendantIds.has(t.id)) continue
      descendantIds.add(t.id)
      queue.push(t.id)
    }
  }
  return descendantIds
}

function findLastDescendantIndex(tasks: GanttTask[], taskId: string): number {
  const descendantIds = collectDescendantIds(tasks, taskId)
  let lastIndex = tasks.findIndex((t) => t.id === taskId)
  for (let i = lastIndex + 1; i < tasks.length; i += 1) {
    if (descendantIds.has(tasks[i].id)) lastIndex = i
  }
  return lastIndex
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
  const [theme, cycleTheme] = useTheme()
  const [workbook, setWorkbook] = useState<GanttWorkbookState>(() =>
    getInitialWorkbook(initialTasks),
  )
  const [renamingSheetId, setRenamingSheetId] = useState<string | null>(null)
  const [renameDraft, setRenameDraft] = useState('')
  const [draggingTaskId, setDraggingTaskId] = useState<string | null>(null)
  const [dragOverTaskId, setDragOverTaskId] = useState<string | null>(null)
  const [dragOverFooter, setDragOverFooter] = useState(false)
  const [includeDayColumnsInExport, setIncludeDayColumnsInExport] = useState(false)
  const [draggingBarTaskId, setDraggingBarTaskId] = useState<string | null>(null)
  const dragBarStateRef = useRef<{
    taskId: string
    mode: BarDragMode
    startClientX: number
    originalStart: string
    originalEnd: string
  } | null>(null)

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

  const setTasksState = useCallback(
    (updater: SetStateAction<GanttTask[]>) => {
      const id = activeSheetId
      setWorkbook((w) => ({
        ...w,
        sheets: w.sheets.map((s) => {
          if (s.id !== id) return s
          const next = typeof updater === 'function' ? updater(s.tasks) : updater
          return { ...s, tasks: next }
        }),
      }))
    },
    [activeSheetId],
  )

  const updateTask = useCallback(
    (taskId: string, patch: Partial<GanttTask>) => {
      setTasksState((prev) =>
        prev.map((t) => {
          if (t.id !== taskId) return t
          const next = { ...t, ...patch }
          if (parseISOToUtcMs(next.end) < parseISOToUtcMs(next.start)) next.end = next.start
          return next
        }),
      )
    },
    [setTasksState],
  )

  function removeTask(taskId: string) {
    setTasksState((prev) => {
      const descendants = collectDescendantIds(prev, taskId)
      return prev.filter((t) => t.id !== taskId && !descendants.has(t.id))
    })
  }

  function addSubtask(parentTask: GanttTask) {
    setTasksState((prev) => {
      if (!prev.some((x) => x.id === parentTask.id)) return prev
      const insertAfter = findLastDescendantIndex(prev, parentTask.id)
      const newTask = createTask({
        name: parentTask.name.trim()
          ? `${parentTask.name.trim()} - subtask`
          : 'New subtask',
        parentId: parentTask.id,
        start: parentTask.start,
        end: parentTask.end,
        progress: 0,
      })
      const next = [...prev]
      next.splice(insertAfter + 1, 0, newTask)
      return next
    })
  }

  function applyVisibleRangeShortcut(range: { start: string; end: string }) {
    patchActiveSheet({ viewRangeOverride: range })
  }

  function beginBarDrag(
    e: ReactPointerEvent<HTMLElement>,
    task: GanttTask,
    mode: BarDragMode,
  ) {
    if (e.button !== 0) return
    dragBarStateRef.current = {
      taskId: task.id,
      mode,
      startClientX: e.clientX,
      originalStart: task.start,
      originalEnd: task.end,
    }
    setDraggingBarTaskId(task.id)
    e.currentTarget.setPointerCapture(e.pointerId)
    e.preventDefault()
  }

  useEffect(() => {
    function onPointerMove(e: PointerEvent) {
      const drag = dragBarStateRef.current
      if (!drag) return
      const deltaX = e.clientX - drag.startClientX
      const dayShift = Math.round(deltaX / DAY_PX)
      if (drag.mode === 'move') {
        updateTask(drag.taskId, {
          start: addDaysISO(drag.originalStart, dayShift),
          end: addDaysISO(drag.originalEnd, dayShift),
        })
        return
      }
      if (drag.mode === 'resize-start') {
        const candidateStart = addDaysISO(drag.originalStart, dayShift)
        updateTask(drag.taskId, {
          start:
            parseISOToUtcMs(candidateStart) > parseISOToUtcMs(drag.originalEnd)
              ? drag.originalEnd
              : candidateStart,
        })
        return
      }
      const candidateEnd = addDaysISO(drag.originalEnd, dayShift)
      updateTask(drag.taskId, {
        end:
          parseISOToUtcMs(candidateEnd) < parseISOToUtcMs(drag.originalStart)
            ? drag.originalStart
            : candidateEnd,
      })
    }

    function onPointerUp() {
      dragBarStateRef.current = null
      setDraggingBarTaskId(null)
    }

    window.addEventListener('pointermove', onPointerMove)
    window.addEventListener('pointerup', onPointerUp)
    return () => {
      window.removeEventListener('pointermove', onPointerMove)
      window.removeEventListener('pointerup', onPointerUp)
    }
  }, [updateTask])

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
      {/* ── Sticky app bar ── */}
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
            <button
              type="button"
              className="gantt-btn gantt-btn--ghost gantt-btn--theme"
              onClick={cycleTheme}
              aria-label={`Theme: ${theme}. Click to cycle`}
              title={`Theme: ${theme}`}
            >
              {theme === 'dark' ? (
                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                  <path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z" />
                </svg>
              ) : theme === 'light' ? (
                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                  <circle cx="12" cy="12" r="5" />
                  <line x1="12" y1="1" x2="12" y2="3" />
                  <line x1="12" y1="21" x2="12" y2="23" />
                  <line x1="4.22" y1="4.22" x2="5.64" y2="5.64" />
                  <line x1="18.36" y1="18.36" x2="19.78" y2="19.78" />
                  <line x1="1" y1="12" x2="3" y2="12" />
                  <line x1="21" y1="12" x2="23" y2="12" />
                  <line x1="4.22" y1="19.78" x2="5.64" y2="18.36" />
                  <line x1="18.36" y1="5.64" x2="19.78" y2="4.22" />
                </svg>
              ) : (
                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                  <circle cx="12" cy="12" r="10" />
                  <path d="M12 2a10 10 0 0 1 0 20" />
                </svg>
              )}
            </button>
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
                  const depth = countAncestorDepth(tasks, t)
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
                        <div
                          className="gantt-task-name-cell"
                          style={{ '--task-indent-level': String(depth) } as CSSProperties}
                        >
                          {depth === 0
                            ? <span className="gantt-task-swatch" aria-hidden="true" />
                            : <span className="gantt-subtask-indicator" aria-hidden="true">↳</span>
                          }
                          <input
                            className="gantt-input gantt-input--task-name"
                            aria-label={`Name for ${t.name}`}
                            value={t.name}
                            onChange={(e) => updateTask(t.id, { name: e.target.value })}
                          />
                          <button
                            type="button"
                            className="gantt-btn gantt-btn--subtask gantt-btn--subtask-inline"
                            aria-label={`Add subtask under ${t.name}`}
                            title="Add subtask"
                            onClick={() => addSubtask(t)}
                          >
                            + Sub
                          </button>
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
            <span className="gantt-section__hint">Drag bars to move · drag edges to resize</span>
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
                    {tasks.map((t, idx) => {
                      const depth = countAncestorDepth(tasks, t)
                      return (
                        <div
                          key={t.id}
                          className={`gantt-chart__task-name${draggingTaskId === t.id ? ' gantt-chart__task-name--dragging' : ''}${dragOverTaskId === t.id ? ' gantt-chart__task-name--drop-target' : ''}`}
                          style={
                            {
                              '--task-color': taskColor(idx),
                              '--task-indent-level': String(depth),
                            } as CSSProperties
                          }
                          title={t.name}
                          draggable
                          onDragStart={(e) => {
                            e.dataTransfer.setData(DND_TASK_MIME, t.id)
                            e.dataTransfer.setData('text/plain', t.id)
                            e.dataTransfer.effectAllowed = 'move'
                            setDraggingTaskId(t.id)
                          }}
                          onDragEnd={endDragSession}
                          onDragOver={(e) => {
                            if (!draggingTaskId) return
                            e.preventDefault()
                            e.dataTransfer.dropEffect = 'move'
                            setDragOverTaskId(t.id)
                          }}
                          onDrop={(e) => {
                            e.preventDefault()
                            const dragId = readDragTaskId(e.dataTransfer)
                            if (!dragId || dragId === t.id) { endDragSession(); return }
                            setTasksState((prev) => reorderTasks(prev, dragId, t.id))
                            endDragSession()
                          }}
                        >
                          <span className="gantt-chart__task-name-dot" aria-hidden="true" />
                          <span className="gantt-chart__task-name-text">
                            {depth > 0 ? '↳ ' : ''}{t.name.trim() || 'Untitled'}
                          </span>
                        </div>
                      )
                    })}
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
                          className={`gantt-chart__track${dragOverTaskId === t.id ? ' gantt-chart__track--drop-target' : ''}`}
                          style={
                            {
                              width: totalDays * DAY_PX,
                              '--task-color': taskColor(idx),
                            } as CSSProperties
                          }
                          onDragOver={(e) => {
                            if (!draggingTaskId) return
                            e.preventDefault()
                            e.dataTransfer.dropEffect = 'move'
                            setDragOverTaskId(t.id)
                          }}
                          onDrop={(e) => {
                            e.preventDefault()
                            const dragId = readDragTaskId(e.dataTransfer)
                            if (!dragId || dragId === t.id) { endDragSession(); return }
                            setTasksState((prev) => reorderTasks(prev, dragId, t.id))
                            endDragSession()
                          }}
                        >
                          {intersects ? (
                            <div
                              className={`gantt-chart__bar${draggingBarTaskId === t.id ? ' gantt-chart__bar--dragging' : ''}`}
                              style={{ left: leftPx, width: widthPx }}
                              onPointerDown={(e) => beginBarDrag(e, t, 'move')}
                            >
                              <button
                                type="button"
                                className="gantt-chart__bar-resize gantt-chart__bar-resize--start"
                                aria-label={`Resize start date for ${t.name.trim() || 'Untitled task'}`}
                                onPointerDown={(e) => {
                                  e.stopPropagation()
                                  beginBarDrag(e, t, 'resize-start')
                                }}
                              />
                              <span
                                className="gantt-chart__bar-fill"
                                style={{ width: `${t.progress}%` }}
                              />
                              <button
                                type="button"
                                className="gantt-chart__bar-resize gantt-chart__bar-resize--end"
                                aria-label={`Resize end date for ${t.name.trim() || 'Untitled task'}`}
                                onPointerDown={(e) => {
                                  e.stopPropagation()
                                  beginBarDrag(e, t, 'resize-end')
                                }}
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
