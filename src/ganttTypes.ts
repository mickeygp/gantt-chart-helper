export type GanttTask = {
  id: string
  name: string
  /** Optional parent task id for subtasks. */
  parentId?: string
  /** YYYY-MM-DD */
  start: string
  /** YYYY-MM-DD */
  end: string
  /** 0–100 */
  progress: number
}

export function createTask(partial?: Partial<Omit<GanttTask, 'id'>>): GanttTask {
  return {
    id: crypto.randomUUID(),
    name: partial?.name ?? 'New task',
    parentId: partial?.parentId,
    start: partial?.start ?? todayISO(),
    end: partial?.end ?? partial?.start ?? todayISO(),
    progress: partial?.progress ?? 0,
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

export function hasSubtasks(task: GanttTask, tasks: GanttTask[]): boolean {
  return tasks.some((t) => t.parentId === task.id)
}
