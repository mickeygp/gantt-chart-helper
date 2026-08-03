import { parseISOToUtcMs } from './ganttDates'

export type GanttTask = {
  id: string
  name: string
  /** Optional parent task id for subtasks. */
  parentId?: string
  /** YYYY-MM-DD, or null when the date has not been decided yet. */
  start: string | null
  /** YYYY-MM-DD, or null when the date has not been decided yet. */
  end: string | null
  /** 0–100 */
  progress: number
  /** When true, this task's subtree is hidden in the task list and timeline. */
  collapsed?: boolean
  /** Whether the task gets a bar on the timeline. Absent means plotted. */
  plotted?: boolean
  /**
   * Ids of tasks this one follows (finish-to-start). Moving a predecessor
   * ripples this task by the same number of days, so whatever gap you left
   * between them is preserved. See ganttDeps.ts.
   */
  deps?: string[]
}

export function createTask(partial?: Partial<Omit<GanttTask, 'id'>>): GanttTask {
  return {
    id: crypto.randomUUID(),
    name: partial?.name ?? 'New task',
    parentId: partial?.parentId,
    start: partial?.start ?? null,
    end: partial?.end ?? partial?.start ?? null,
    progress: partial?.progress ?? 0,
    collapsed: partial?.collapsed,
    plotted: partial?.plotted,
    deps: partial?.deps?.length ? [...partial.deps] : undefined,
  }
}

export function todayISO(): string {
  const d = new Date()
  const y = d.getFullYear()
  const m = String(d.getMonth() + 1).padStart(2, '0')
  const day = String(d.getDate()).padStart(2, '0')
  return `${y}-${m}-${day}`
}

function clampPct(n: number): number {
  if (!Number.isFinite(n)) return 0
  return Math.min(100, Math.max(0, n))
}

/**
 * Returns the simple average of immediate children's effective progress
 * (recursive). Tasks with no children fall back to their own stored value.
 */
export function getEffectiveProgress(task: GanttTask, tasks: GanttTask[]): number {
  const children = tasks.filter((t) => t.parentId === task.id)
  if (children.length === 0) return clampPct(task.progress)
  const sum = children.reduce((acc, c) => acc + getEffectiveProgress(c, tasks), 0)
  return clampPct(sum / children.length)
}

export function isPlotted(task: GanttTask): boolean {
  return task.plotted !== false
}

// ── Date resolution ─────────────────────────────────────────────────────────
//
// Dates are optional, so a task's position on the timeline can come from three
// places: its own fields, a roll-up of its subtasks, or nowhere at all (the
// task is unscheduled and gets no bar).

export type TaskRange = {
  start: string
  end: string
  /** True when at least one endpoint was rolled up from subtasks. */
  derived: boolean
}

function earlier(a: string, b: string): string {
  return parseISOToUtcMs(a) <= parseISOToUtcMs(b) ? a : b
}

function later(a: string, b: string): string {
  return parseISOToUtcMs(a) >= parseISOToUtcMs(b) ? a : b
}

/**
 * Resolves every task's timeline range in one pass, keyed by task id. A null
 * entry means the task has no dates of its own and no dated subtasks.
 */
export function resolveTaskRanges(tasks: GanttTask[]): Map<string, TaskRange | null> {
  const childrenByParent = new Map<string, GanttTask[]>()
  for (const t of tasks) {
    if (!t.parentId) continue
    const siblings = childrenByParent.get(t.parentId)
    if (siblings) siblings.push(t)
    else childrenByParent.set(t.parentId, [t])
  }

  const resolved = new Map<string, TaskRange | null>()
  const inProgress = new Set<string>()

  function resolve(task: GanttTask): TaskRange | null {
    const cached = resolved.get(task.id)
    if (cached !== undefined) return cached
    // Guards against a corrupted parent chain forming a cycle.
    if (inProgress.has(task.id)) return null
    inProgress.add(task.id)

    let rollupStart: string | null = null
    let rollupEnd: string | null = null
    for (const child of childrenByParent.get(task.id) ?? []) {
      const childRange = resolve(child)
      if (!childRange) continue
      rollupStart = rollupStart ? earlier(rollupStart, childRange.start) : childRange.start
      rollupEnd = rollupEnd ? later(rollupEnd, childRange.end) : childRange.end
    }

    const start = task.start ?? rollupStart ?? task.end
    const end = task.end ?? rollupEnd ?? task.start
    let range: TaskRange | null = null
    if (start && end) {
      range = {
        start: earlier(start, end),
        end: later(start, end),
        derived: (!task.start && rollupStart !== null) || (!task.end && rollupEnd !== null),
      }
    }

    inProgress.delete(task.id)
    resolved.set(task.id, range)
    return range
  }

  for (const t of tasks) resolve(t)
  return resolved
}

/** Union of the given tasks' resolved ranges, or null when none are scheduled. */
export function unionRanges(
  tasks: GanttTask[],
  ranges: Map<string, TaskRange | null>,
): { start: string; end: string } | null {
  let start: string | null = null
  let end: string | null = null
  for (const t of tasks) {
    const r = ranges.get(t.id)
    if (!r) continue
    start = start ? earlier(start, r.start) : r.start
    end = end ? later(end, r.end) : r.end
  }
  return start && end ? { start, end } : null
}

/** Tasks whose ancestors are all expanded, in document order. */
export function getVisibleTasks(tasks: GanttTask[]): GanttTask[] {
  const byId = new Map(tasks.map((t) => [t.id, t]))
  return tasks.filter((t) => {
    let parentId = t.parentId
    let depth = 0
    while (parentId && depth < 64) {
      const parent = byId.get(parentId)
      if (!parent) break
      if (parent.collapsed) return false
      parentId = parent.parentId
      depth += 1
    }
    return true
  })
}
