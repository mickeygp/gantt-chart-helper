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
