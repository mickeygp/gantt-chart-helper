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
import { DragDropProvider, useDroppable, type DragEndEvent } from '@dnd-kit/react'
import { useSortable } from '@dnd-kit/react/sortable'

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
import { addDaysISO, daysInclusive, diffDays, parseISOToUtcMs } from './ganttDates'
import {
  collectDownstreamIds,
  eligibleDepCandidates,
  pruneDanglingDeps,
  rippleFrom,
  validDeps,
} from './ganttDeps'
import { buildMonthSpans, buildWeekSpans, type WeekLabelFormat } from './ganttTimeline'
import { loadGanttWorkbook, saveGanttWorkbook } from './ganttLocalCache'
import { fetchSharedSheet, publishSharedSheet } from './ganttCloud'
import { useGanttCloud, type CloudStatus } from './useGanttCloud'
import {
  createSheet,
  duplicateSheet,
  nextSheetLabel,
  type GanttSheetState,
  type GanttWorkbookState,
} from './ganttSheet'
import {
  createTask,
  getEffectiveProgress,
  getVisibleTasks,
  isPlotted,
  resolveTaskRanges,
  unionRanges,
  type GanttTask,
  type TaskRange,
  todayISO,
} from './ganttTypes'
import './GanttBuilder.css'

const DEFAULT_DAY_PX = 26
const MIN_DAY_PX = 2
const MAX_DAY_PX = 60
const WEEK_LABEL_FORMATS: readonly WeekLabelFormat[] = ['iso', 'month', 'date']
const WEEK_LABEL_FORMAT_OPTIONS: { value: WeekLabelFormat; label: string }[] = [
  { value: 'iso', label: 'ISO week (W22)' },
  { value: 'month', label: 'Week of month (W1)' },
  { value: 'date', label: 'Date of Monday' },
]
const ZOOM_STEP = 4
const DEFAULT_LABEL_COL_W = 180
const MIN_LABEL_COL_W = 100
const MAX_LABEL_COL_W = 420
const FOOTER_DROP_ID = '__footer__'
const TIMELINE_LABEL_PREFIX = 'tl-'
const TIMELINE_TRACK_PREFIX = 'tltrack-'
const TABLE_SORT_GROUP = 'task-table'
const TIMELINE_SORT_GROUP = 'task-timeline'
type BarDragMode = 'move' | 'resize-start' | 'resize-end'

const CLOUD_LABELS: Record<CloudStatus, string> = {
  off: 'Local only',
  'signed-out': 'Sign in',
  syncing: 'Syncing…',
  conflict: 'Sync paused',
  saving: 'Saving…',
  saved: 'Synced',
  error: 'Sync error',
}

function timelineLabelId(taskId: string) { return TIMELINE_LABEL_PREFIX + taskId }
function timelineTrackId(taskId: string) { return TIMELINE_TRACK_PREFIX + taskId }
function normalizeTaskId(id: string): string {
  if (id.startsWith(TIMELINE_TRACK_PREFIX)) return id.slice(TIMELINE_TRACK_PREFIX.length)
  if (id.startsWith(TIMELINE_LABEL_PREFIX)) return id.slice(TIMELINE_LABEL_PREFIX.length)
  return id
}

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

// ── Shared row affordances ──────────────────────────────────────────────────

interface CollapseToggleProps {
  task: GanttTask
  subtaskCount: number
  onToggle: (task: GanttTask) => void
  size?: 'sm'
}

/**
 * Disclosure triangle for parent tasks. Leaf tasks render an inert spacer so
 * names stay aligned down the column.
 */
function CollapseToggle({ task, subtaskCount, onToggle, size }: CollapseToggleProps) {
  const cls = `gantt-collapse-toggle${size === 'sm' ? ' gantt-collapse-toggle--sm' : ''}`
  if (subtaskCount === 0) {
    return <span className={`${cls} gantt-collapse-toggle--empty`} aria-hidden="true" />
  }
  const collapsed = task.collapsed === true
  const label = task.name.trim() || 'Untitled task'
  return (
    <button
      type="button"
      className={cls}
      aria-expanded={!collapsed}
      aria-label={`${collapsed ? 'Expand' : 'Collapse'} ${subtaskCount} subtask${subtaskCount === 1 ? '' : 's'} of ${label}`}
      title={collapsed ? `Expand (${subtaskCount})` : `Collapse (${subtaskCount})`}
      onClick={() => onToggle(task)}
    >
      <svg viewBox="0 0 12 12" aria-hidden="true" focusable="false">
        <path d="M4 2.5 L8 6 L4 9.5" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" />
      </svg>
    </button>
  )
}

interface PlotAllCheckboxProps {
  allPlotted: boolean
  nonePlotted: boolean
  disabled: boolean
  onToggle: () => void
}

/** Header checkbox that plots or un-plots every task, mixed state included. */
function PlotAllCheckbox({ allPlotted, nonePlotted, disabled, onToggle }: PlotAllCheckboxProps) {
  const ref = useRef<HTMLInputElement | null>(null)
  const indeterminate = !allPlotted && !nonePlotted
  useEffect(() => {
    if (ref.current) ref.current.indeterminate = indeterminate
  }, [indeterminate])
  return (
    <input
      ref={ref}
      type="checkbox"
      className="gantt-checkbox"
      checked={allPlotted}
      disabled={disabled}
      aria-label="Plot all tasks on the timeline"
      title={allPlotted ? 'Un-plot all tasks' : 'Plot all tasks'}
      onChange={onToggle}
    />
  )
}

interface DepsCellProps {
  t: GanttTask
  tasks: GanttTask[]
  addDep: (taskId: string, depId: string) => void
  removeDep: (taskId: string, depId: string) => void
}

/**
 * Predecessor chips plus a picker. The picker only offers tasks that cannot
 * close a loop, so there is no invalid state to warn about after the fact.
 */
function DepsCell({ t, tasks, addDep, removeDep }: DepsCellProps) {
  const byId = useMemo(() => new Map(tasks.map((x) => [x.id, x])), [tasks])
  const deps = useMemo(() => validDeps(t, tasks), [t, tasks])
  const candidates = useMemo(() => eligibleDepCandidates(tasks, t.id), [tasks, t.id])
  const taskLabel = t.name.trim() || 'Untitled task'

  return (
    <div className="gantt-deps-cell">
      {deps.map((depId) => {
        const dep = byId.get(depId)
        const depLabel = dep?.name.trim() || 'Untitled task'
        return (
          <span className="gantt-dep-chip" key={depId} title={`After: ${depLabel}`}>
            <span className="gantt-dep-chip__text">{depLabel}</span>
            <button
              type="button"
              className="gantt-dep-chip__remove"
              aria-label={`Unlink ${taskLabel} from ${depLabel}`}
              onClick={() => removeDep(t.id, depId)}
            >
              ×
            </button>
          </span>
        )
      })}
      <select
        className="gantt-input gantt-input-select gantt-deps-cell__picker"
        value=""
        aria-label={`Add a task that ${taskLabel} follows`}
        title={
          candidates.length
            ? 'Link this task to run after another one'
            : 'No other task can be linked without creating a loop'
        }
        disabled={candidates.length === 0}
        onChange={(e) => {
          if (e.target.value) addDep(t.id, e.target.value)
        }}
      >
        <option value="">{deps.length ? '+ link' : 'After…'}</option>
        {candidates.map((c) => (
          <option key={c.id} value={c.id}>
            {c.name.trim() || 'Untitled task'}
          </option>
        ))}
      </select>
    </div>
  )
}

/** Explains where a blank date field gets its timeline position from. */
function dateFieldHint(range: TaskRange | null, side: 'start' | 'end'): string {
  if (!range) return 'No date set — add a date here or on a subtask to plot this task'
  const fallback = side === 'start' ? range.start : range.end
  return range.derived
    ? `No date set — using ${fallback} from subtasks`
    : `No date set — using ${fallback} from the other date`
}

// ── dnd-kit sub-components ──────────────────────────────────────────────────

interface SortableTaskRowProps {
  t: GanttTask
  index: number
  colorIndex: number
  tasks: GanttTask[]
  range: TaskRange | null
  subtaskCount: number
  updateTask: (id: string, patch: Partial<GanttTask>) => void
  updateTaskEnd: (id: string, end: string | null) => void
  removeTask: (id: string) => void
  addSubtask: (task: GanttTask) => void
  toggleCollapse: (task: GanttTask) => void
  togglePlotted: (task: GanttTask) => void
  addDep: (taskId: string, depId: string) => void
  removeDep: (taskId: string, depId: string) => void
}

function SortableTaskRow({
  t,
  index,
  colorIndex,
  tasks,
  range,
  subtaskCount,
  updateTask,
  updateTaskEnd,
  removeTask,
  addSubtask,
  toggleCollapse,
  togglePlotted,
  addDep,
  removeDep,
}: SortableTaskRowProps) {
  const { ref, handleRef, isDragging, isDropTarget } = useSortable({
    id: t.id,
    index,
    group: TABLE_SORT_GROUP,
  })
  const depth = countAncestorDepth(tasks, t)
  const color = taskColor(colorIndex)
  const isParent = subtaskCount > 0
  const plotted = isPlotted(t)
  const effectiveProgress = isParent
    ? Math.round(getEffectiveProgress(t, tasks))
    : t.progress
  const taskLabel = t.name.trim() || 'Untitled task'
  return (
    <tr
      ref={ref as unknown as React.RefCallback<HTMLTableRowElement>}
      data-task-id={t.id}
      className={`gantt-table__task-row${isDragging ? ' gantt-table__task-row--dragging' : ''}${isDropTarget ? ' gantt-table__task-row--drop-target' : ''}${plotted ? '' : ' gantt-table__task-row--unplotted'}`}
      style={{ '--task-color': color } as CSSProperties}
    >
      <td className="gantt-table__cell-plot">
        <input
          type="checkbox"
          className="gantt-checkbox"
          checked={plotted}
          aria-label={`Plot ${taskLabel} on the timeline`}
          title={plotted ? 'Plotted on the timeline' : 'Hidden from the timeline'}
          onChange={() => togglePlotted(t)}
        />
      </td>
      <td className="gantt-table__cell-drag">
        <button
          ref={handleRef as unknown as React.RefCallback<HTMLButtonElement>}
          type="button"
          className="gantt-drag-handle"
          aria-label={`Drag to reorder row: ${t.name.trim() || 'Untitled task'}`}
          title="Drag to reorder"
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
          <CollapseToggle task={t} subtaskCount={subtaskCount} onToggle={toggleCollapse} />
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
          {t.collapsed && subtaskCount > 0 ? (
            <span className="gantt-subtask-count" title={`${subtaskCount} hidden subtask${subtaskCount === 1 ? '' : 's'}`}>
              {subtaskCount}
            </span>
          ) : null}
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
          className={`gantt-input gantt-input--date${t.start === null ? ' gantt-input--date-unset' : ''}`}
          type="date"
          aria-label={`Start date for ${taskLabel} (optional)`}
          title={t.start === null ? dateFieldHint(range, 'start') : undefined}
          value={t.start ?? ''}
          onChange={(e) => {
            const start = e.target.value || null
            const clampEnd =
              start !== null && t.end !== null && parseISOToUtcMs(t.end) < parseISOToUtcMs(start)
            updateTask(t.id, { start, end: clampEnd ? start : t.end })
          }}
        />
      </td>
      <td>
        <input
          className={`gantt-input gantt-input--date${t.end === null ? ' gantt-input--date-unset' : ''}`}
          type="date"
          aria-label={`End date for ${taskLabel} (optional)`}
          title={t.end === null ? dateFieldHint(range, 'end') : undefined}
          value={t.end ?? ''}
          min={t.start ?? undefined}
          onChange={(e) => updateTaskEnd(t.id, e.target.value || null)}
        />
      </td>
      <td>
        <DepsCell t={t} tasks={tasks} addDep={addDep} removeDep={removeDep} />
      </td>
      <td className="gantt-num">
        {range ? (
          <span className={range.derived ? 'gantt-num__derived' : undefined} title={range.derived ? `Rolled up from subtasks: ${range.start} → ${range.end}` : undefined}>
            {daysInclusive(range.start, range.end)}
          </span>
        ) : (
          <span className="gantt-num__unset" title="No dates yet">—</span>
        )}
      </td>
      <td>
        <div className="gantt-progress-cell">
          <input
            className="gantt-input gantt-input--narrow"
            type="number"
            min={0}
            max={100}
            aria-label={`Progress for ${t.name}`}
            value={effectiveProgress}
            disabled={isParent}
            readOnly={isParent}
            title={isParent ? 'Computed from subtasks (average)' : undefined}
            onFocus={(e) => e.currentTarget.select()}
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
              style={{ width: `${effectiveProgress}%` }}
            />
          </div>
        </div>
      </td>
      <td>
        <button
          type="button"
          className="gantt-btn gantt-btn--delete"
          aria-label={`Remove ${t.name}`}
          title="Remove task (⌘/Ctrl + Backspace)"
          onClick={() => removeTask(t.id)}
        >
          ✕
        </button>
      </td>
    </tr>
  )
}

interface DroppableFooterProps {
  setTasksState: (updater: SetStateAction<GanttTask[]>) => void
  isDragging: boolean
}

function DroppableFooter({ setTasksState, isDragging }: DroppableFooterProps) {
  const { ref, isDropTarget } = useDroppable({ id: FOOTER_DROP_ID })
  return (
    <tfoot>
      <tr
        ref={ref as unknown as React.RefCallback<HTMLTableRowElement>}
        className={`gantt-table__foot-row${isDropTarget ? ' gantt-table__foot-row--drop' : ''}`}
      >
        <td colSpan={9}>
          <div className="gantt-table__foot-inner">
            <button
              type="button"
              className="gantt-btn gantt-btn--secondary gantt-table__add-btn"
              onClick={() => setTasksState((prev) => [...prev, createTask()])}
            >
              + Add task
            </button>
            {isDragging ? (
              <span className="gantt-table__drop-hint" aria-live="polite">
                Drop here to move to bottom
              </span>
            ) : null}
          </div>
        </td>
      </tr>
    </tfoot>
  )
}

interface SortableTimelineTaskNameProps {
  t: GanttTask
  index: number
  colorIndex: number
  tasks: GanttTask[]
  subtaskCount: number
  toggleCollapse: (task: GanttTask) => void
}

function SortableTimelineTaskName({
  t,
  index,
  colorIndex,
  tasks,
  subtaskCount,
  toggleCollapse,
}: SortableTimelineTaskNameProps) {
  const { ref, handleRef, isDragging, isDropTarget } = useSortable({
    id: timelineLabelId(t.id),
    index,
    group: TIMELINE_SORT_GROUP,
  })
  const depth = countAncestorDepth(tasks, t)
  return (
    <div
      ref={ref as unknown as React.RefCallback<HTMLDivElement>}
      className={`gantt-chart__task-name${isDragging ? ' gantt-chart__task-name--dragging' : ''}${isDropTarget ? ' gantt-chart__task-name--drop-target' : ''}`}
      style={
        {
          '--task-color': taskColor(colorIndex),
          '--task-indent-level': String(depth),
        } as CSSProperties
      }
      title={t.name}
    >
      <button
        ref={handleRef as unknown as React.RefCallback<HTMLButtonElement>}
        type="button"
        className="gantt-drag-handle gantt-drag-handle--sm"
        aria-label={`Drag to reorder: ${t.name.trim() || 'Untitled task'}`}
        title="Drag to reorder"
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
      <CollapseToggle task={t} subtaskCount={subtaskCount} onToggle={toggleCollapse} size="sm" />
      <span className="gantt-chart__task-name-dot" aria-hidden="true" />
      <span className="gantt-chart__task-name-text">
        {depth > 0 ? '↳ ' : ''}{t.name.trim() || 'Untitled'}
      </span>
    </div>
  )
}

interface DroppableTimelineTrackProps {
  t: GanttTask
  colorIndex: number
  draggingBarTaskId: string | null
  beginBarDrag: (e: ReactPointerEvent<HTMLElement>, task: GanttTask, mode: BarDragMode) => void
  totalDays: number
  leftPx: number
  widthPx: number
  intersects: boolean
  /** True when the bar's dates roll up from subtasks, so it can't be dragged. */
  isSummary: boolean
  /** True when neither the task nor its subtasks have any date. */
  isUnscheduled: boolean
  dayPx: number
  progressPct: number
}

function DroppableTimelineTrack({
  t,
  colorIndex,
  draggingBarTaskId,
  beginBarDrag,
  totalDays,
  leftPx,
  widthPx,
  intersects,
  isSummary,
  isUnscheduled,
  dayPx,
  progressPct,
}: DroppableTimelineTrackProps) {
  const { ref, isDropTarget } = useDroppable({ id: timelineTrackId(t.id) })
  const taskLabel = t.name.trim() || 'Untitled task'
  return (
    <div
      ref={ref as unknown as React.RefCallback<HTMLDivElement>}
      className={`gantt-chart__track${isDropTarget ? ' gantt-chart__track--drop-target' : ''}`}
      style={
        {
          width: totalDays * dayPx,
          '--task-color': taskColor(colorIndex),
        } as CSSProperties
      }
    >
      {isUnscheduled ? (
        <span className="gantt-chart__unscheduled">No dates yet</span>
      ) : intersects && isSummary ? (
        <div
          className="gantt-chart__bar gantt-chart__bar--summary"
          style={{ left: leftPx, width: widthPx }}
          title={`${taskLabel}: dates rolled up from subtasks`}
        >
          <span className="gantt-chart__bar-fill" style={{ width: `${progressPct}%` }} />
        </div>
      ) : intersects ? (
        <div
          className={`gantt-chart__bar${draggingBarTaskId === t.id ? ' gantt-chart__bar--dragging' : ''}`}
          style={{ left: leftPx, width: widthPx }}
          onPointerDown={(e) => beginBarDrag(e, t, 'move')}
        >
          <button
            type="button"
            className="gantt-chart__bar-resize gantt-chart__bar-resize--start"
            aria-label={`Resize start date for ${taskLabel}`}
            onPointerDown={(e) => {
              e.stopPropagation()
              beginBarDrag(e, t, 'resize-start')
            }}
          />
          <span
            className="gantt-chart__bar-fill"
            style={{ width: `${progressPct}%` }}
          />
          <button
            type="button"
            className="gantt-chart__bar-resize gantt-chart__bar-resize--end"
            aria-label={`Resize end date for ${taskLabel}`}
            onPointerDown={(e) => {
              e.stopPropagation()
              beginBarDrag(e, t, 'resize-end')
            }}
          />
        </div>
      ) : null}
    </div>
  )
}

type TrackGeometry = { top: number; pitch: number }

/**
 * Where the track rows actually sit inside the timeline column.
 *
 * The header and row heights come from CSS variables and shift at the mobile
 * breakpoint, and neither element uses border-box, so the arrow overlay
 * measures the rendered rows instead of recomputing them from constants that
 * would drift the moment the stylesheet changes.
 */
function useTrackGeometry(
  colRef: React.RefObject<HTMLDivElement | null>,
  deps: unknown[],
): TrackGeometry | null {
  const [geometry, setGeometry] = useState<TrackGeometry | null>(null)

  useEffect(() => {
    const col = colRef.current
    if (!col) return

    function measure() {
      const el = colRef.current
      if (!el) return
      const rows = el.querySelectorAll<HTMLElement>('.gantt-chart__track')
      if (rows.length === 0) {
        setGeometry(null)
        return
      }
      const first = rows[0]
      const pitch =
        rows.length > 1 ? rows[1].offsetTop - first.offsetTop : first.offsetHeight
      setGeometry((prev) =>
        prev && prev.top === first.offsetTop && prev.pitch === pitch
          ? prev
          : { top: first.offsetTop, pitch },
      )
    }

    measure()
    const observer = new ResizeObserver(measure)
    observer.observe(col)
    return () => observer.disconnect()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [colRef, ...deps])

  return geometry
}

interface DependencyArrowsProps {
  /** Plotted, visible tasks in row order. */
  rows: GanttTask[]
  ranges: Map<string, TaskRange | null>
  rangeStart: string
  rangeEnd: string
  dayPx: number
  width: number
  geometry: TrackGeometry
}

/**
 * Elbow connectors from each predecessor's finish to its successor's start.
 *
 * Only links where both ends are on screen are drawn — an arrow to a collapsed,
 * un-plotted, or off-range task would point at nothing.
 */
function DependencyArrows({
  rows,
  ranges,
  rangeStart,
  rangeEnd,
  dayPx,
  width,
  geometry,
}: DependencyArrowsProps) {
  const rowIndex = useMemo(
    () => new Map(rows.map((t, i) => [t.id, i])),
    [rows],
  )

  const paths = useMemo(() => {
    const out: { key: string; d: string; x: number; y: number }[] = []
    const stub = 10
    const r0 = parseISOToUtcMs(rangeStart)
    const r1 = parseISOToUtcMs(rangeEnd)

    function xForDayStart(iso: string): number {
      return (diffDays(rangeStart, iso) * dayPx)
    }
    function centerY(index: number): number {
      return geometry.top + index * geometry.pitch + geometry.pitch / 2
    }
    function onScreen(range: TaskRange | null): boolean {
      if (!range) return false
      return parseISOToUtcMs(range.end) >= r0 && parseISOToUtcMs(range.start) <= r1
    }

    for (const task of rows) {
      const toIdx = rowIndex.get(task.id)
      const toRange = ranges.get(task.id) ?? null
      if (toIdx === undefined || !onScreen(toRange) || !toRange) continue

      for (const depId of task.deps ?? []) {
        const fromIdx = rowIndex.get(depId)
        const fromRange = ranges.get(depId) ?? null
        if (fromIdx === undefined || !onScreen(fromRange) || !fromRange) continue

        // Finish edge of the predecessor, start edge of the successor.
        const x1 = xForDayStart(fromRange.end) + dayPx
        const y1 = centerY(fromIdx)
        const x2 = xForDayStart(toRange.start)
        const y2 = centerY(toIdx)

        let d: string
        if (x2 - x1 >= stub * 2) {
          // Room for a normal elbow: across, down, into the successor.
          d = `M ${x1} ${y1} H ${x2 - stub} V ${y2} H ${x2}`
        } else if (x2 >= x1) {
          // Back-to-back tasks, the usual case: a clean vertical drop. Routing
          // an elbow through this sliver would draw a visible zigzag.
          d = `M ${x1} ${y1} V ${y2} H ${x2}`
        } else {
          // The successor genuinely starts before its predecessor finishes, so
          // the connector doubles back through the gutter between the rows.
          const gutter = y1 + (y2 >= y1 ? geometry.pitch / 2 : -geometry.pitch / 2)
          d = `M ${x1} ${y1} h ${stub} V ${gutter} H ${x2 - stub} V ${y2} H ${x2}`
        }

        out.push({ key: `${depId}->${task.id}`, d, x: x2, y: y2 })
      }
    }
    return out
  }, [rows, rowIndex, ranges, rangeStart, rangeEnd, dayPx, geometry])

  if (paths.length === 0) return null

  return (
    <svg
      className="gantt-chart__deps"
      width={width}
      height={geometry.top + rows.length * geometry.pitch}
      viewBox={`0 0 ${width} ${geometry.top + rows.length * geometry.pitch}`}
      aria-hidden="true"
      focusable="false"
    >
      {paths.map((p) => (
        <g key={p.key}>
          <path className="gantt-chart__dep-line" d={p.d} />
          <path
            className="gantt-chart__dep-head"
            d={`M ${p.x} ${p.y} l -5 -3.5 l 0 7 z`}
          />
        </g>
      ))}
    </svg>
  )
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
  const [includeDayColumnsInExport, setIncludeDayColumnsInExport] = useState(false)
  const [draggingBarTaskId, setDraggingBarTaskId] = useState<string | null>(null)
  const [isRowDragging, setIsRowDragging] = useState(false)
  const [showShortcuts, setShowShortcuts] = useState(false)
  const [showAccount, setShowAccount] = useState(false)
  const [sharing, setSharing] = useState(false)
  const [notice, setNotice] = useState<{ kind: 'info' | 'error'; text: string } | null>(
    null,
  )
  const [dayPx, setDayPx] = useState<number>(() => {
    const stored = Number(localStorage.getItem('gantt-day-px'))
    if (Number.isFinite(stored) && stored >= MIN_DAY_PX && stored <= MAX_DAY_PX) return stored
    return DEFAULT_DAY_PX
  })
  const [labelColWidth, setLabelColWidth] = useState<number>(() => {
    const stored = Number(localStorage.getItem('gantt-label-col-w'))
    if (Number.isFinite(stored) && stored >= MIN_LABEL_COL_W && stored <= MAX_LABEL_COL_W) return stored
    return DEFAULT_LABEL_COL_W
  })
  const [weekLabelFormat, setWeekLabelFormat] = useState<WeekLabelFormat>(() => {
    const stored = localStorage.getItem('gantt-week-label-format')
    return (WEEK_LABEL_FORMATS as readonly string[]).includes(stored ?? '')
      ? (stored as WeekLabelFormat)
      : 'iso'
  })
  const labelResizeRef = useRef<{ startClientX: number; startWidth: number } | null>(null)
  const [isLabelResizing, setIsLabelResizing] = useState(false)
  const scrollContainerRef = useRef<HTMLDivElement | null>(null)
  const timelineColRef = useRef<HTMLDivElement | null>(null)
  const fitTimelineRef = useRef<() => void>(() => {})

  useEffect(() => {
    localStorage.setItem('gantt-day-px', String(dayPx))
  }, [dayPx])

  useEffect(() => {
    localStorage.setItem('gantt-label-col-w', String(labelColWidth))
  }, [labelColWidth])

  useEffect(() => {
    localStorage.setItem('gantt-week-label-format', weekLabelFormat)
  }, [weekLabelFormat])

  useEffect(() => {
    function onMove(e: PointerEvent) {
      const r = labelResizeRef.current
      if (!r) return
      const dx = e.clientX - r.startClientX
      const next = Math.max(MIN_LABEL_COL_W, Math.min(MAX_LABEL_COL_W, r.startWidth + dx))
      setLabelColWidth(next)
    }
    function onUp() {
      if (!labelResizeRef.current) return
      labelResizeRef.current = null
      setIsLabelResizing(false)
    }
    window.addEventListener('pointermove', onMove)
    window.addEventListener('pointerup', onUp)
    return () => {
      window.removeEventListener('pointermove', onMove)
      window.removeEventListener('pointerup', onUp)
    }
  }, [])

  function beginLabelColResize(e: ReactPointerEvent<HTMLElement>) {
    if (e.button !== 0) return
    labelResizeRef.current = { startClientX: e.clientX, startWidth: labelColWidth }
    setIsLabelResizing(true)
    e.currentTarget.setPointerCapture(e.pointerId)
    e.preventDefault()
  }

  const zoomOut = useCallback(
    () => setDayPx((d) => Math.max(MIN_DAY_PX, d - ZOOM_STEP)),
    [],
  )
  const zoomIn = useCallback(
    () => setDayPx((d) => Math.min(MAX_DAY_PX, d + ZOOM_STEP)),
    [],
  )
  const resetZoom = useCallback(() => setDayPx(DEFAULT_DAY_PX), [])
  const fitTimeline = useCallback(() => fitTimelineRef.current(), [])
  const dragBarStateRef = useRef<{
    taskId: string
    mode: BarDragMode
    startClientX: number
    originalStart: string
    originalEnd: string
    /** Which endpoints the task actually stores, so a move keeps blanks blank. */
    hadStart: boolean
    hadEnd: boolean
    /**
     * Pre-drag dates of every task linked downstream. Each frame recomputes
     * their positions from this snapshot rather than nudging the previous
     * frame, so dragging back and forth lands exactly where it started.
     */
    downstream: Map<string, { start: string | null; end: string | null }>
  } | null>(null)

  function handleDragEnd({ operation }: DragEndEvent) {
    setIsRowDragging(false)
    if (operation.canceled) return
    const rawSourceId = operation.source?.id as string | undefined
    const rawTargetId = operation.target?.id as string | undefined
    if (!rawSourceId) return
    const sourceId = normalizeTaskId(rawSourceId)
    if (rawTargetId === FOOTER_DROP_ID) {
      setTasksState((prev) => moveTaskToEnd(prev, sourceId))
    } else if (rawTargetId) {
      const targetId = normalizeTaskId(rawTargetId)
      if (targetId !== sourceId) {
        setTasksState((prev) => reorderTasks(prev, sourceId, targetId))
      }
    }
  }

  const { sheets, activeSheetId } = workbook

  const activeSheet = useMemo(
    () => sheets.find((s) => s.id === activeSheetId) ?? sheets[0],
    [sheets, activeSheetId],
  )

  const { projectName, tasks, viewRangeOverride } = activeSheet

  const taskRanges = useMemo(() => resolveTaskRanges(tasks), [tasks])

  const subtaskCounts = useMemo(() => {
    const counts = new Map<string, number>()
    for (const t of tasks) {
      if (!t.parentId) continue
      counts.set(t.parentId, (counts.get(t.parentId) ?? 0) + 1)
    }
    return counts
  }, [tasks])

  /** Index in the full task list, so bar colors stay put as rows collapse. */
  const colorIndexById = useMemo(() => {
    const map = new Map<string, number>()
    tasks.forEach((t, i) => map.set(t.id, i))
    return map
  }, [tasks])

  /** Rows the task table shows: everything except subtrees of collapsed parents. */
  const visibleTasks = useMemo(() => getVisibleTasks(tasks), [tasks])

  /** Rows the timeline draws: visible rows that are also ticked for plotting. */
  const timelineTasks = useMemo(() => visibleTasks.filter(isPlotted), [visibleTasks])

  const collapsibleTasks = useMemo(
    () => tasks.filter((t) => subtaskCounts.has(t.id)),
    [tasks, subtaskCounts],
  )
  const hiddenRowCount = tasks.length - visibleTasks.length
  const plottedCount = useMemo(() => tasks.filter(isPlotted).length, [tasks])
  const allPlotted = tasks.length > 0 && plottedCount === tasks.length
  const nonePlotted = plottedCount === 0
  const canExpandAny = collapsibleTasks.some((t) => t.collapsed)
  const canCollapseAny = collapsibleTasks.some((t) => !t.collapsed)

  const autoRange = useMemo(
    () => unionRanges(timelineTasks, taskRanges),
    [timelineTasks, taskRanges],
  )

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
      totalWidth: days.length * dayPx,
      monthSpans: buildMonthSpans(days),
      weekSpans: buildWeekSpans(days, weekLabelFormat),
    }
  }, [effectiveRange, dayPx, weekLabelFormat])

  const trackGeometry = useTrackGeometry(timelineColRef, [
    timelineTasks.length,
    dayPx,
    labelColWidth,
  ])

  const todayOffsetPx = useMemo(() => {
    const today = todayISO()
    if (today < effectiveRange.start || today > effectiveRange.end) return null
    const offsetDays = Math.round(
      (parseISOToUtcMs(today) - parseISOToUtcMs(effectiveRange.start)) / 86_400_000,
    )
    return offsetDays * dayPx + dayPx / 2
  }, [effectiveRange, dayPx])

  const weekOffset = useMemo(() => {
    const ms = parseISOToUtcMs(effectiveRange.start)
    if (Number.isNaN(ms)) return 0
    return (new Date(ms).getUTCDay() + 6) % 7
  }, [effectiveRange.start])

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

  const cloud = useGanttCloud(workbook, setWorkbook, initialTasks !== undefined)

  // ── Share links ───────────────────────────────────────────────────────────

  /** Import runs once per page load even under StrictMode's double effect. */
  const shareImportedRef = useRef(false)

  useEffect(() => {
    if (initialTasks !== undefined || shareImportedRef.current) return
    const shareId = new URLSearchParams(window.location.search).get('share')
    if (!shareId) return
    shareImportedRef.current = true

    void fetchSharedSheet(shareId).then((res) => {
      // Drop the token from the address bar either way, so a reload does not
      // re-import the same sheet a second time.
      const url = new URL(window.location.href)
      url.searchParams.delete('share')
      window.history.replaceState(null, '', url.toString())

      if (!res.ok) {
        setNotice({ kind: 'error', text: res.error })
        return
      }
      // A shared sheet arrives with the sender's ids; re-key it so it cannot
      // collide with a sheet already open here.
      setWorkbook((w) => {
        const imported = duplicateSheet(res.data, w.sheets)
        return {
          activeSheetId: imported.id,
          sheets: [...w.sheets, imported],
        }
      })
      setNotice({
        kind: 'info',
        text: `Imported "${res.data.sheetName}" as a new sheet.`,
      })
    })
  }, [initialTasks])

  async function shareActiveSheet() {
    const sheet = workbook.sheets.find((s) => s.id === activeSheetId)
    if (!sheet) return
    setSharing(true)
    const res = await publishSharedSheet(sheet)
    setSharing(false)
    if (!res.ok) {
      setNotice({ kind: 'error', text: res.error })
      return
    }
    const url = new URL(window.location.href)
    url.search = ''
    url.searchParams.set('share', res.data)
    const link = url.toString()
    try {
      await navigator.clipboard.writeText(link)
      setNotice({ kind: 'info', text: 'Share link copied to your clipboard.' })
    } catch {
      // Clipboard access can be denied; the link still has to reach the user.
      setNotice({ kind: 'info', text: `Share link: ${link}` })
    }
  }

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
          if (
            next.start !== null &&
            next.end !== null &&
            parseISOToUtcMs(next.end) < parseISOToUtcMs(next.start)
          ) {
            next.end = next.start
          }
          return next
        }),
      )
    },
    [setTasksState],
  )

  /**
   * Setting a finish date by hand ripples exactly like dragging the bar's right
   * edge — the two ways of moving a task should not disagree.
   */
  const updateTaskEnd = useCallback(
    (taskId: string, nextEnd: string | null) => {
      setTasksState((prev) => {
        const current = prev.find((t) => t.id === taskId)
        if (!current) return prev
        // Same clamp as dragging the right edge: the finish never precedes the start.
        const end =
          nextEnd !== null &&
          current.start !== null &&
          parseISOToUtcMs(nextEnd) < parseISOToUtcMs(current.start)
            ? current.start
            : nextEnd
        const delta = current.end && end ? diffDays(current.end, end) : 0
        const rippled = delta !== 0 ? rippleFrom(prev, taskId, delta) : prev
        return rippled.map((t) => (t.id === taskId ? { ...t, end } : t))
      })
    },
    [setTasksState],
  )

  const addDep = useCallback(
    (taskId: string, depId: string) => {
      setTasksState((prev) =>
        prev.map((t) =>
          t.id === taskId
            ? { ...t, deps: [...new Set([...(t.deps ?? []), depId])] }
            : t,
        ),
      )
    },
    [setTasksState],
  )

  const removeDep = useCallback(
    (taskId: string, depId: string) => {
      setTasksState((prev) =>
        prev.map((t) => {
          if (t.id !== taskId) return t
          const kept = (t.deps ?? []).filter((d) => d !== depId)
          return { ...t, deps: kept.length ? kept : undefined }
        }),
      )
    },
    [setTasksState],
  )

  const toggleCollapse = useCallback(
    (task: GanttTask) => {
      setTasksState((prev) =>
        prev.map((t) => (t.id === task.id ? { ...t, collapsed: !t.collapsed } : t)),
      )
    },
    [setTasksState],
  )

  const setAllCollapsed = useCallback(
    (collapsed: boolean) => {
      setTasksState((prev) => {
        const parentIds = new Set(prev.map((t) => t.parentId).filter(Boolean) as string[])
        return prev.map((t) =>
          parentIds.has(t.id) ? { ...t, collapsed: collapsed || undefined } : t,
        )
      })
    },
    [setTasksState],
  )

  /** Ticking a parent cascades to its whole subtree, matching how rows read. */
  const togglePlotted = useCallback(
    (task: GanttTask) => {
      setTasksState((prev) => {
        const next = !isPlotted(task)
        const affected = collectDescendantIds(prev, task.id)
        affected.add(task.id)
        return prev.map((t) =>
          affected.has(t.id) ? { ...t, plotted: next ? undefined : false } : t,
        )
      })
    },
    [setTasksState],
  )

  const setAllPlotted = useCallback(
    (plotted: boolean) => {
      setTasksState((prev) => prev.map((t) => ({ ...t, plotted: plotted ? undefined : false })))
    },
    [setTasksState],
  )

  function removeTask(taskId: string) {
    setTasksState((prev) => {
      const descendants = collectDescendantIds(prev, taskId)
      const kept = prev.filter((t) => t.id !== taskId && !descendants.has(t.id))
      // Links into the deleted subtree would otherwise linger as dead ids.
      return pruneDanglingDeps(kept)
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
        plotted: parentTask.plotted,
      })
      const next = [...prev]
      next.splice(insertAfter + 1, 0, newTask)
      // Adding a subtask to a collapsed parent would hide it, so expand first.
      return next.map((t) => (t.id === parentTask.id ? { ...t, collapsed: undefined } : t))
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
    const range = taskRanges.get(task.id)
    // Rolled-up bars follow their subtasks, so they aren't dragged directly.
    if (!range || range.derived) return
    const downstream = new Map<string, { start: string | null; end: string | null }>()
    for (const id of collectDownstreamIds(tasks, task.id)) {
      const t = tasks.find((x) => x.id === id)
      if (t) downstream.set(id, { start: t.start, end: t.end })
    }
    dragBarStateRef.current = {
      taskId: task.id,
      mode,
      startClientX: e.clientX,
      originalStart: range.start,
      originalEnd: range.end,
      hadStart: task.start !== null,
      hadEnd: task.end !== null,
      downstream,
    }
    setDraggingBarTaskId(task.id)
    e.currentTarget.setPointerCapture(e.pointerId)
    e.preventDefault()
  }

  useEffect(() => {
    function onPointerMove(e: PointerEvent) {
      const drag = dragBarStateRef.current
      if (!drag) return
      const dayShift = Math.round((e.clientX - drag.startClientX) / dayPx)

      // Only a change to the finish date pushes successors under finish-to-start,
      // so dragging the left edge (which changes duration, not finish) stays put.
      let rippleShift = 0
      if (drag.mode === 'move') {
        rippleShift = dayShift
      } else if (drag.mode === 'resize-end') {
        const candidate = addDaysISO(drag.originalEnd, dayShift)
        const clamped =
          parseISOToUtcMs(candidate) < parseISOToUtcMs(drag.originalStart)
            ? drag.originalStart
            : candidate
        // The clamp at the start date caps how far successors actually move.
        rippleShift = diffDays(drag.originalEnd, clamped)
      }

      setTasksState((prev) =>
        prev.map((t) => {
          if (t.id === drag.taskId) {
            if (drag.mode === 'move') {
              return {
                ...t,
                start: drag.hadStart ? addDaysISO(drag.originalStart, dayShift) : t.start,
                end: drag.hadEnd ? addDaysISO(drag.originalEnd, dayShift) : t.end,
              }
            }
            if (drag.mode === 'resize-start') {
              const candidate = addDaysISO(drag.originalStart, dayShift)
              return {
                ...t,
                start:
                  parseISOToUtcMs(candidate) > parseISOToUtcMs(drag.originalEnd)
                    ? drag.originalEnd
                    : candidate,
              }
            }
            return { ...t, end: addDaysISO(drag.originalEnd, rippleShift) }
          }

          const original = drag.downstream.get(t.id)
          if (!original) return t
          // Recomputed from the snapshot on every frame, including a shift of
          // zero: skipping that case would strand successors wherever the
          // previous frame left them when the drag returns to its start.
          const start =
            original.start === null ? null : addDaysISO(original.start, rippleShift)
          const end =
            original.end === null ? null : addDaysISO(original.end, rippleShift)
          if (start === t.start && end === t.end) return t
          return { ...t, start, end }
        }),
      )
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
  }, [setTasksState, dayPx])

  fitTimelineRef.current = () => {
    const el = scrollContainerRef.current
    if (!el) return
    const numDays = timeline.days.length
    if (numDays <= 0) return
    const available = el.clientWidth - labelColWidth
    if (available <= 0) return
    const raw = available / numDays
    const computed = Math.max(MIN_DAY_PX, Math.min(MAX_DAY_PX, raw))
    setDayPx(Math.round(computed * 100) / 100)
  }

  const shortcutHandlerRef = useRef<(e: KeyboardEvent) => void>(() => {})
  shortcutHandlerRef.current = (e: KeyboardEvent) => {
    if (e.key === 'Escape') {
      if (showShortcuts) { e.preventDefault(); setShowShortcuts(false); return }
      if (showAccount) { e.preventDefault(); setShowAccount(false); return }
      if (notice) { e.preventDefault(); setNotice(null); return }
      if (renamingSheetId !== null) { e.preventDefault(); cancelRename(); return }
      return
    }

    if ((e.metaKey || e.ctrlKey) && !e.altKey && !e.shiftKey && e.key === 'Backspace') {
      const target = e.target instanceof HTMLElement ? e.target : null
      const row = target?.closest('[data-task-id]') as HTMLElement | null
      const taskId = row?.dataset.taskId
      if (taskId) {
        e.preventDefault()
        removeTask(taskId)
      }
      return
    }

    if (e.metaKey || e.ctrlKey || e.altKey) return

    const target = e.target
    const inTextField =
      target instanceof HTMLElement &&
      (target.tagName === 'INPUT' ||
        target.tagName === 'TEXTAREA' ||
        target.tagName === 'SELECT' ||
        target.isContentEditable)
    if (inTextField) return

    if (e.key === '?') {
      e.preventDefault()
      setShowShortcuts((v) => !v)
      return
    }

    if (e.shiftKey && e.key === 'C') {
      e.preventDefault()
      addSheet()
      return
    }

    if (e.shiftKey && e.key === 'D') {
      e.preventDefault()
      duplicateSheetById(activeSheetId)
      return
    }

    if (e.key === '+' || e.key === '=') {
      e.preventDefault()
      zoomIn()
      return
    }
    if (e.key === '-') {
      e.preventDefault()
      zoomOut()
      return
    }
    if (e.key === '0') {
      e.preventDefault()
      fitTimeline()
      return
    }
    if (e.key === '[') {
      e.preventDefault()
      setAllCollapsed(true)
      return
    }
    if (e.key === ']') {
      e.preventDefault()
      setAllCollapsed(false)
      return
    }

    if (e.shiftKey) return

    switch (e.key) {
      case 'c':
        e.preventDefault()
        setTasksState((prev) => [...prev, createTask()])
        return
      case 'e':
        e.preventDefault()
        exportGanttToXlsx(tasks, {
          projectName,
          visibleRange: effectiveRange,
          includeDayColumns: includeDayColumnsInExport,
          weekLabelFormat,
        })
        return
      case 't':
        e.preventDefault()
        cycleTheme()
        return
      case 'f':
        e.preventDefault()
        patchActiveSheet({ viewRangeOverride: null })
        return
      case '1':
      case '2':
      case '3':
      case '4':
      case '5': {
        const idx = Number(e.key) - 1
        const shortcut = visibleRangeShortcuts[idx]
        if (shortcut) {
          e.preventDefault()
          applyVisibleRangeShortcut(shortcut.getRange())
        }
        return
      }
    }
  }

  useEffect(() => {
    function onKeyDown(e: KeyboardEvent) {
      shortcutHandlerRef.current(e)
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [])

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

  /**
   * Copies a sheet and opens it, so "plan A vs plan B" starts from the real
   * plan instead of a re-typed one. The copy lands directly to the right of
   * its source rather than at the end of the tab strip.
   */
  function duplicateSheetById(id: string) {
    setRenamingSheetId(null)
    setWorkbook((w) => {
      const source = w.sheets.find((s) => s.id === id)
      if (!source) return w
      const copy = duplicateSheet(source, w.sheets)
      const sheets = [...w.sheets]
      sheets.splice(w.sheets.indexOf(source) + 1, 0, copy)
      return { activeSheetId: copy.id, sheets }
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
          gridTemplateColumns: `repeat(${timeline.days.length}, ${dayPx}px)`,
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
    <DragDropProvider
      onDragStart={() => setIsRowDragging(true)}
      onDragEnd={handleDragEnd}
    >
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
            {cloud.configured && (
              <>
                <button
                  type="button"
                  className={`gantt-btn gantt-btn--ghost gantt-cloud-pill gantt-cloud-pill--${cloud.status}`}
                  onClick={() => setShowAccount((v) => !v)}
                  aria-expanded={showAccount}
                  title={
                    cloud.email
                      ? `Signed in as ${cloud.email}`
                      : 'Sign in with Google to sync this workbook across browsers'
                  }
                >
                  <span className="gantt-cloud-pill__dot" aria-hidden="true" />
                  <span className="gantt-btn__label">{CLOUD_LABELS[cloud.status]}</span>
                </button>
                <button
                  type="button"
                  className="gantt-btn gantt-btn--ghost"
                  onClick={() => void shareActiveSheet()}
                  disabled={sharing || !cloud.email}
                  title={
                    cloud.email
                      ? 'Publish a snapshot of this sheet and copy a link to it'
                      : 'Sign in to create share links'
                  }
                >
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                    <circle cx="18" cy="5" r="3" />
                    <circle cx="6" cy="12" r="3" />
                    <circle cx="18" cy="19" r="3" />
                    <line x1="8.6" y1="10.5" x2="15.4" y2="6.5" />
                    <line x1="8.6" y1="13.5" x2="15.4" y2="17.5" />
                  </svg>
                  <span className="gantt-btn__label">{sharing ? 'Sharing…' : 'Share'}</span>
                </button>
              </>
            )}
            <button
              type="button"
              className="gantt-btn gantt-btn--ghost gantt-btn--theme"
              onClick={() => setShowShortcuts(true)}
              aria-label="Show keyboard shortcuts"
              title="Keyboard shortcuts (?)"
            >
              <span className="gantt-shortcuts-hint-glyph" aria-hidden="true">?</span>
            </button>
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
                  weekLabelFormat,
                })
              }
            >
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                <polyline points="7 10 12 15 17 10" />
                <line x1="12" y1="15" x2="12" y2="3" />
              </svg>
              <span className="gantt-btn__label">Export XLSX</span>
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
            {hiddenRowCount > 0 && (
              <span className="gantt-section__hint">{hiddenRowCount} hidden</span>
            )}
            {collapsibleTasks.length > 0 && (
              <div className="gantt-section__actions" role="group" aria-label="Expand and collapse">
                <button
                  type="button"
                  className="gantt-btn gantt-btn--ghost gantt-btn--with-kbd"
                  title="Collapse every task that has subtasks — press ["
                  disabled={!canCollapseAny}
                  onClick={() => setAllCollapsed(true)}
                >
                  <span>Collapse all</span>
                  <kbd className="gantt-btn__kbd" aria-hidden="true">[</kbd>
                </button>
                <button
                  type="button"
                  className="gantt-btn gantt-btn--ghost gantt-btn--with-kbd"
                  title="Expand every task that has subtasks — press ]"
                  disabled={!canExpandAny}
                  onClick={() => setAllCollapsed(false)}
                >
                  <span>Expand all</span>
                  <kbd className="gantt-btn__kbd" aria-hidden="true">]</kbd>
                </button>
              </div>
            )}
          </div>

          <section className="gantt-table-wrap" aria-label="Task list">
            <table className="gantt-table">
              <thead>
                <tr>
                  <th scope="col" className="gantt-table__th-plot">
                    <PlotAllCheckbox
                      allPlotted={allPlotted}
                      nonePlotted={nonePlotted}
                      disabled={tasks.length === 0}
                      onToggle={() => setAllPlotted(!allPlotted)}
                    />
                  </th>
                  <th scope="col" className="gantt-table__th-drag">
                    <span className="visually-hidden">Reorder</span>
                  </th>
                  <th scope="col">Task</th>
                  <th scope="col">Start</th>
                  <th scope="col">End</th>
                  <th scope="col" title="Tasks this one runs after. Moving them moves this task too.">
                    Depends on
                  </th>
                  <th scope="col">Days</th>
                  <th scope="col">Progress</th>
                  <th scope="col">
                    <span className="visually-hidden">Remove</span>
                  </th>
                </tr>
              </thead>
              <tbody>
                {visibleTasks.map((t, idx) => (
                  <SortableTaskRow
                    key={t.id}
                    t={t}
                    index={idx}
                    colorIndex={colorIndexById.get(t.id) ?? idx}
                    tasks={tasks}
                    range={taskRanges.get(t.id) ?? null}
                    subtaskCount={subtaskCounts.get(t.id) ?? 0}
                    updateTask={updateTask}
                    updateTaskEnd={updateTaskEnd}
                    removeTask={removeTask}
                    addSubtask={addSubtask}
                    toggleCollapse={toggleCollapse}
                    togglePlotted={togglePlotted}
                    addDep={addDep}
                    removeDep={removeDep}
                  />
                ))}
              </tbody>
              <DroppableFooter setTasksState={setTasksState} isDragging={isRowDragging} />
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
                {visibleRangeShortcuts.map((shortcut, i) => (
                  <button
                    key={shortcut.key}
                    type="button"
                    className="gantt-btn gantt-btn--ghost gantt-btn--with-kbd"
                    title={`${shortcut.title} — press ${i + 1}`}
                    onClick={() => applyVisibleRangeShortcut(shortcut.getRange())}
                  >
                    <span>{shortcut.label}</span>
                    <kbd className="gantt-btn__kbd" aria-hidden="true">{i + 1}</kbd>
                  </button>
                ))}
              </div>
              <div className="gantt-chart__zoom" role="group" aria-label="Timeline zoom">
                <button
                  type="button"
                  className="gantt-chart__zoom-btn"
                  aria-label="Zoom out"
                  title="Zoom out (−)"
                  onClick={zoomOut}
                  disabled={dayPx <= MIN_DAY_PX}
                >
                  −
                </button>
                <button
                  type="button"
                  className="gantt-chart__zoom-value"
                  title="Reset zoom"
                  aria-label={`Day width ${dayPx} pixels — click to reset`}
                  onClick={resetZoom}
                >
                  {Number.isInteger(dayPx) ? dayPx : dayPx.toFixed(1)}px
                </button>
                <button
                  type="button"
                  className="gantt-chart__zoom-btn"
                  aria-label="Zoom in"
                  title="Zoom in (+)"
                  onClick={zoomIn}
                  disabled={dayPx >= MAX_DAY_PX}
                >
                  +
                </button>
              </div>
              <label className="gantt-chart__week-label-select">
                <span className="visually-hidden">Week label format</span>
                <select
                  className="gantt-input gantt-input-select"
                  value={weekLabelFormat}
                  onChange={(e) => setWeekLabelFormat(e.target.value as WeekLabelFormat)}
                  aria-label="Week label format"
                  title="Week label format"
                >
                  {WEEK_LABEL_FORMAT_OPTIONS.map((opt) => (
                    <option key={opt.value} value={opt.value}>
                      {opt.label}
                    </option>
                  ))}
                </select>
              </label>
              <button
                type="button"
                className="gantt-btn gantt-btn--ghost gantt-btn--with-kbd gantt-chart__zoom-fit"
                title="Auto-zoom so the entire visible range fits — press 0"
                onClick={fitTimeline}
              >
                <span>Fit width</span>
                <kbd className="gantt-btn__kbd" aria-hidden="true">0</kbd>
              </button>
              <button
                type="button"
                className="gantt-btn gantt-btn--secondary gantt-chart__range-fit gantt-btn--with-kbd"
                title={
                  autoRange
                    ? 'Snap the visible range to your tasks (clears manual dates) — press F'
                    : 'Clear manual range and use the default window until you add tasks — press F'
                }
                onClick={() => patchActiveSheet({ viewRangeOverride: null })}
              >
                <span>Fit all tasks</span>
                <kbd className="gantt-btn__kbd" aria-hidden="true">F</kbd>
              </button>
            </div>

            {timeline.days.length === 0 ? (
              <p className="gantt-chart__empty">Set a valid date range.</p>
            ) : timelineTasks.length === 0 ? (
              <>
                <div className="gantt-chart__scroll" ref={scrollContainerRef}>
                  <div
                    className="gantt-chart__pan"
                    style={{
                      '--gantt-day-px': `${dayPx}px`,
                      '--gantt-week-offset': String(weekOffset),
                    } as CSSProperties}
                  >
                    <div
                      className="gantt-chart__label-col"
                      style={{ flex: `0 0 ${labelColWidth}px` }}
                    >
                      <div className="gantt-chart__label-header-spacer" aria-hidden="true" />
                      <div
                        className={`gantt-chart__label-col-resizer${isLabelResizing ? ' gantt-chart__label-col-resizer--active' : ''}`}
                        role="separator"
                        aria-orientation="vertical"
                        aria-label="Resize task name column"
                        title="Drag to resize"
                        onPointerDown={beginLabelColResize}
                        onDoubleClick={() => setLabelColWidth(DEFAULT_LABEL_COL_W)}
                      />
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
                  {tasks.length === 0
                    ? 'Add a task to see bars on the timeline.'
                    : 'No tasks are plotted. Tick a task in the checkbox column to add it here.'}
                </p>
              </>
            ) : (
              <div className="gantt-chart__scroll" ref={scrollContainerRef}>
                <div
                  className="gantt-chart__pan"
                  style={{
                    '--gantt-day-px': `${dayPx}px`,
                    '--gantt-week-offset': String(weekOffset),
                  } as CSSProperties}
                >
                  <div
                    className="gantt-chart__label-col"
                    style={{ flex: `0 0 ${labelColWidth}px` }}
                  >
                    <div className="gantt-chart__label-header-spacer" aria-hidden="true" />
                    {timelineTasks.map((t, idx) => (
                      <SortableTimelineTaskName
                        key={t.id}
                        t={t}
                        index={idx}
                        colorIndex={colorIndexById.get(t.id) ?? idx}
                        tasks={tasks}
                        subtaskCount={subtaskCounts.get(t.id) ?? 0}
                        toggleCollapse={toggleCollapse}
                      />
                    ))}
                    <div
                      className={`gantt-chart__label-col-resizer${isLabelResizing ? ' gantt-chart__label-col-resizer--active' : ''}`}
                      role="separator"
                      aria-orientation="vertical"
                      aria-label="Resize task name column"
                      title="Drag to resize · double-click to reset"
                      onPointerDown={beginLabelColResize}
                      onDoubleClick={() => setLabelColWidth(DEFAULT_LABEL_COL_W)}
                    />
                  </div>
                  <div
                    className="gantt-chart__timeline-col"
                    ref={timelineColRef}
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
                    {timelineTasks.map((t, idx) => {
                      const rangeStart = timeline.days[0] ?? effectiveRange.start
                      const rangeEnd = timeline.days[timeline.days.length - 1] ?? effectiveRange.end
                      const dayMs = 86_400_000
                      const spanMs = parseISOToUtcMs(rangeEnd) - parseISOToUtcMs(rangeStart)
                      const totalDays = spanMs >= 0 ? Math.floor(spanMs / dayMs) + 1 : 1

                      const taskRange = taskRanges.get(t.id) ?? null
                      const r0 = parseISOToUtcMs(rangeStart)
                      const r1 = parseISOToUtcMs(rangeEnd)
                      const t0 = taskRange ? parseISOToUtcMs(taskRange.start) : NaN
                      const t1 = taskRange ? parseISOToUtcMs(taskRange.end) : NaN
                      const intersects = taskRange !== null && t1 >= r0 && t0 <= r1

                      let leftPx = 0
                      let widthPx = 0
                      if (intersects && taskRange) {
                        const visStart = t0 < r0 ? rangeStart : taskRange.start
                        const visEnd = t1 > r1 ? rangeEnd : taskRange.end
                        const offsetDays = Math.round(
                          (parseISOToUtcMs(visStart) - r0) / dayMs,
                        )
                        leftPx = offsetDays * dayPx
                        widthPx = Math.max(daysInclusive(visStart, visEnd) * dayPx, 4)
                      }

                      return (
                        <DroppableTimelineTrack
                          key={t.id}
                          t={t}
                          colorIndex={colorIndexById.get(t.id) ?? idx}
                          draggingBarTaskId={draggingBarTaskId}
                          beginBarDrag={beginBarDrag}
                          totalDays={totalDays}
                          leftPx={leftPx}
                          widthPx={widthPx}
                          intersects={intersects}
                          isSummary={taskRange?.derived === true}
                          isUnscheduled={taskRange === null}
                          dayPx={dayPx}
                          progressPct={getEffectiveProgress(t, tasks)}
                        />
                      )
                    })}
                    {trackGeometry && (
                      <DependencyArrows
                        rows={timelineTasks}
                        ranges={taskRanges}
                        rangeStart={timeline.days[0] ?? effectiveRange.start}
                        rangeEnd={
                          timeline.days[timeline.days.length - 1] ?? effectiveRange.end
                        }
                        dayPx={dayPx}
                        width={timeline.totalWidth}
                        geometry={trackGeometry}
                      />
                    )}
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
                  {isRenaming ? null : (
                    <button
                      type="button"
                      className="gantt-tabs__duplicate"
                      aria-label={`Duplicate ${s.sheetName}`}
                      title="Duplicate sheet"
                      tabIndex={isActive ? 0 : -1}
                      onClick={(e) => { e.stopPropagation(); duplicateSheetById(s.id) }}
                    >
                      <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                        <rect x="9" y="9" width="12" height="12" rx="2" />
                        <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1" />
                      </svg>
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

      {showAccount && cloud.configured && (
        <div
          className="gantt-shortcuts-overlay"
          role="dialog"
          aria-modal="true"
          aria-label="Cloud sync"
          onClick={() => setShowAccount(false)}
        >
          <div className="gantt-shortcuts-card gantt-account-card" onClick={(e) => e.stopPropagation()}>
            <div className="gantt-shortcuts-card__head">
              <h2 className="gantt-shortcuts-card__title">Cloud sync</h2>
              <button
                type="button"
                className="gantt-shortcuts-card__close"
                aria-label="Close"
                onClick={() => setShowAccount(false)}
              >
                ×
              </button>
            </div>
            <div className="gantt-account-card__body">
              {cloud.email ? (
                <>
                  <p className="gantt-account-card__line">
                    Signed in as <strong>{cloud.email}</strong>
                  </p>
                  <p className="gantt-account-card__hint">
                    Every sheet in this workbook saves to your account, so opening
                    the app in another browser or on another computer brings it
                    back. Status: {CLOUD_LABELS[cloud.status]}.
                  </p>
                  <button
                    type="button"
                    className="gantt-btn gantt-btn--secondary"
                    onClick={() => {
                      void cloud.signOut()
                      setShowAccount(false)
                    }}
                  >
                    Sign out
                  </button>
                </>
              ) : (
                <>
                  <p className="gantt-account-card__hint">
                    Sign in and this workbook follows you between browsers and
                    computers. Until then it stays in this browser only.
                  </p>
                  <button
                    type="button"
                    className="gantt-btn gantt-btn--secondary gantt-google-btn"
                    autoFocus
                    onClick={() => void cloud.signIn()}
                  >
                    <svg width="17" height="17" viewBox="0 0 18 18" aria-hidden="true">
                      <path fill="#4285F4" d="M17.64 9.2c0-.64-.06-1.25-.16-1.84H9v3.48h4.84a4.14 4.14 0 0 1-1.8 2.72v2.26h2.92c1.7-1.57 2.68-3.88 2.68-6.62z" />
                      <path fill="#34A853" d="M9 18c2.43 0 4.47-.8 5.96-2.18l-2.92-2.26c-.8.54-1.84.86-3.04.86-2.34 0-4.32-1.58-5.03-3.7H.96v2.33A9 9 0 0 0 9 18z" />
                      <path fill="#FBBC05" d="M3.97 10.72a5.4 5.4 0 0 1 0-3.44V4.95H.96a9 9 0 0 0 0 8.1l3.01-2.33z" />
                      <path fill="#EA4335" d="M9 3.58c1.32 0 2.5.46 3.44 1.35l2.58-2.58C13.46.9 11.43 0 9 0A9 9 0 0 0 .96 4.95l3.01 2.33C4.68 5.16 6.66 3.58 9 3.58z" />
                    </svg>
                    <span>Continue with Google</span>
                  </button>
                  <p className="gantt-account-card__hint">
                    You'll be sent to Google and returned here. No email is sent,
                    so there is no sign-in link to wait for.
                  </p>
                </>
              )}
              {cloud.error && (
                <p className="gantt-account-card__error" role="alert">
                  {cloud.error}
                  <button
                    type="button"
                    className="gantt-btn gantt-btn--ghost"
                    onClick={cloud.dismissError}
                  >
                    Dismiss
                  </button>
                </p>
              )}
            </div>
          </div>
        </div>
      )}

      {cloud.status === 'conflict' && (
        <div
          className="gantt-shortcuts-overlay"
          role="dialog"
          aria-modal="true"
          aria-label="Choose which version to keep"
        >
          <div className="gantt-shortcuts-card gantt-account-card">
            <div className="gantt-shortcuts-card__head">
              <h2 className="gantt-shortcuts-card__title">Two versions found</h2>
            </div>
            <div className="gantt-account-card__body">
              <p className="gantt-account-card__hint">
                Your account already has a saved workbook, and this browser has a
                different one. Nothing is overwritten until you choose — pick the
                one to keep working in.
              </p>
              <div className="gantt-account-card__choices">
                <button
                  type="button"
                  className="gantt-btn gantt-btn--primary"
                  onClick={() => cloud.resolveConflict('cloud')}
                >
                  Use the saved version
                </button>
                <button
                  type="button"
                  className="gantt-btn gantt-btn--secondary"
                  onClick={() => cloud.resolveConflict('local')}
                >
                  Upload this browser's version
                </button>
                <button
                  type="button"
                  className="gantt-btn gantt-btn--ghost"
                  onClick={() => void cloud.signOut()}
                >
                  Sign out instead
                </button>
              </div>
              <p className="gantt-account-card__hint">
                Want both? Sign out, duplicate the sheets you care about
                (<kbd>Shift</kbd>+<kbd>D</kbd>), then sign back in and upload.
              </p>
            </div>
          </div>
        </div>
      )}

      {notice && (
        <div
          className={`gantt-notice gantt-notice--${notice.kind}`}
          role="status"
          aria-live="polite"
        >
          <span>{notice.text}</span>
          <button
            type="button"
            className="gantt-notice__close"
            aria-label="Dismiss"
            onClick={() => setNotice(null)}
          >
            ×
          </button>
        </div>
      )}

      {showShortcuts && (
        <div
          className="gantt-shortcuts-overlay"
          role="dialog"
          aria-modal="true"
          aria-label="Keyboard shortcuts"
          onClick={() => setShowShortcuts(false)}
        >
          <div
            className="gantt-shortcuts-card"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="gantt-shortcuts-card__head">
              <h2 className="gantt-shortcuts-card__title">Keyboard shortcuts</h2>
              <button
                type="button"
                className="gantt-shortcuts-card__close"
                aria-label="Close"
                onClick={() => setShowShortcuts(false)}
              >
                ×
              </button>
            </div>
            <div className="gantt-shortcuts-card__body">
              <section className="gantt-shortcuts-group">
                <h3 className="gantt-shortcuts-group__title">Tasks</h3>
                <ul className="gantt-shortcuts-list">
                  <li><span className="gantt-shortcuts-keys"><kbd>C</kbd></span><span>Add task</span></li>
                  <li><span className="gantt-shortcuts-keys"><kbd>Shift</kbd><kbd>C</kbd></span><span>New sheet</span></li>
                  <li><span className="gantt-shortcuts-keys"><kbd>Shift</kbd><kbd>D</kbd></span><span>Duplicate current sheet</span></li>
                  <li><span className="gantt-shortcuts-keys"><kbd>⌘</kbd><kbd>⌫</kbd></span><span>Remove focused task (Ctrl+Backspace on Windows)</span></li>
                  <li><span className="gantt-shortcuts-keys"><kbd>[</kbd></span><span>Collapse all subtasks</span></li>
                  <li><span className="gantt-shortcuts-keys"><kbd>]</kbd></span><span>Expand all subtasks</span></li>
                </ul>
              </section>
              <section className="gantt-shortcuts-group">
                <h3 className="gantt-shortcuts-group__title">View</h3>
                <ul className="gantt-shortcuts-list">
                  <li><span className="gantt-shortcuts-keys"><kbd>F</kbd></span><span>Fit all tasks (date range)</span></li>
                  <li><span className="gantt-shortcuts-keys"><kbd>1</kbd></span><span>2 weeks</span></li>
                  <li><span className="gantt-shortcuts-keys"><kbd>2</kbd></span><span>1 month</span></li>
                  <li><span className="gantt-shortcuts-keys"><kbd>3</kbd></span><span>1 quarter</span></li>
                  <li><span className="gantt-shortcuts-keys"><kbd>4</kbd></span><span>1 year</span></li>
                  <li><span className="gantt-shortcuts-keys"><kbd>5</kbd></span><span>This year</span></li>
                </ul>
              </section>
              <section className="gantt-shortcuts-group">
                <h3 className="gantt-shortcuts-group__title">Zoom</h3>
                <ul className="gantt-shortcuts-list">
                  <li><span className="gantt-shortcuts-keys"><kbd>+</kbd></span><span>Zoom in (also <kbd>=</kbd>)</span></li>
                  <li><span className="gantt-shortcuts-keys"><kbd>−</kbd></span><span>Zoom out</span></li>
                  <li><span className="gantt-shortcuts-keys"><kbd>0</kbd></span><span>Fit timeline to width</span></li>
                </ul>
              </section>
              <section className="gantt-shortcuts-group">
                <h3 className="gantt-shortcuts-group__title">App</h3>
                <ul className="gantt-shortcuts-list">
                  <li><span className="gantt-shortcuts-keys"><kbd>E</kbd></span><span>Export XLSX</span></li>
                  <li><span className="gantt-shortcuts-keys"><kbd>T</kbd></span><span>Cycle theme</span></li>
                  <li><span className="gantt-shortcuts-keys"><kbd>?</kbd></span><span>Toggle shortcuts</span></li>
                  <li><span className="gantt-shortcuts-keys"><kbd>Esc</kbd></span><span>Close overlay</span></li>
                </ul>
              </section>
            </div>
            <div className="gantt-shortcuts-card__foot">
              Shortcuts pause while typing in any field.
            </div>
          </div>
        </div>
      )}
    </div>
    </DragDropProvider>
  )
}
