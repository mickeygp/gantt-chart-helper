import { addDaysISO } from './ganttDates'
import type { GanttTask } from './ganttTypes'

/**
 * Finish-to-start links between tasks, with "ripple, keep the gap" semantics:
 * moving a predecessor by N days moves every downstream task by the same N
 * days. Slack you deliberately left in the plan survives the move — nothing is
 * auto-collapsed to butt up against its predecessor.
 *
 * Parent/child rollup complicates who counts as "moved". A task's dates can be
 * derived from its subtasks, so dragging a leaf effectively moves its ancestors
 * too, and shifting a parent has to carry its subtree along. Both directions
 * are handled through `relatedGroup`.
 */

const MAX_WALK = 4096

function childrenByParent(tasks: GanttTask[]): Map<string, GanttTask[]> {
  const map = new Map<string, GanttTask[]>()
  for (const t of tasks) {
    if (!t.parentId) continue
    const siblings = map.get(t.parentId)
    if (siblings) siblings.push(t)
    else map.set(t.parentId, [t])
  }
  return map
}

/** Tasks that list `id` as a predecessor. */
function successorsByDep(tasks: GanttTask[]): Map<string, GanttTask[]> {
  const map = new Map<string, GanttTask[]>()
  for (const t of tasks) {
    for (const depId of t.deps ?? []) {
      const list = map.get(depId)
      if (list) list.push(t)
      else map.set(depId, [t])
    }
  }
  return map
}

function descendantIds(
  children: Map<string, GanttTask[]>,
  rootId: string,
): string[] {
  const out: string[] = []
  const queue = [rootId]
  const seen = new Set([rootId])
  while (queue.length > 0 && out.length < MAX_WALK) {
    const cur = queue.shift()!
    for (const child of children.get(cur) ?? []) {
      if (seen.has(child.id)) continue
      seen.add(child.id)
      out.push(child.id)
      queue.push(child.id)
    }
  }
  return out
}

function ancestorIds(byId: Map<string, GanttTask>, taskId: string): string[] {
  const out: string[] = []
  let parentId = byId.get(taskId)?.parentId
  let depth = 0
  while (parentId && depth < 64) {
    if (out.includes(parentId)) break
    out.push(parentId)
    parentId = byId.get(parentId)?.parentId
    depth += 1
  }
  return out
}

/**
 * A task plus everyone whose timeline position moves with it: its subtree
 * (children follow a parent) and its ancestors (a summary bar follows its
 * children). Used to decide which links a move should trigger.
 */
export function relatedGroup(tasks: GanttTask[], taskId: string): Set<string> {
  const byId = new Map(tasks.map((t) => [t.id, t]))
  const children = childrenByParent(tasks)
  return new Set([
    taskId,
    ...descendantIds(children, taskId),
    ...ancestorIds(byId, taskId),
  ])
}

/**
 * Every task that should ride along when `sourceId` moves, excluding the source
 * and anything already moving with it. Cycles are tolerated: each task is
 * shifted at most once.
 */
export function collectDownstreamIds(
  tasks: GanttTask[],
  sourceId: string,
): Set<string> {
  const byId = new Map(tasks.map((t) => [t.id, t]))
  const children = childrenByParent(tasks)
  const successors = successorsByDep(tasks)

  const anchored = new Set([
    sourceId,
    ...descendantIds(children, sourceId),
    ...ancestorIds(byId, sourceId),
  ])

  const shifted = new Set<string>()
  const queue = [...anchored]
  const visited = new Set(anchored)

  while (queue.length > 0 && shifted.size < MAX_WALK) {
    const cur = queue.shift()!
    for (const successor of successors.get(cur) ?? []) {
      if (anchored.has(successor.id) || shifted.has(successor.id)) continue

      // A successor drags its own subtree with it, but never the anchored set.
      for (const id of [successor.id, ...descendantIds(children, successor.id)]) {
        if (!anchored.has(id)) shifted.add(id)
      }

      // Keep walking: the successor's parents and children move too, so links
      // hanging off any of them fire as well.
      for (const id of [
        successor.id,
        ...descendantIds(children, successor.id),
        ...ancestorIds(byId, successor.id),
      ]) {
        if (visited.has(id)) continue
        visited.add(id)
        queue.push(id)
      }
    }
  }

  return shifted
}

/**
 * Shifts the given tasks by `deltaDays`, leaving blank endpoints blank so a
 * half-scheduled task does not silently gain a date.
 */
export function shiftTasksByDays(
  tasks: GanttTask[],
  ids: ReadonlySet<string>,
  deltaDays: number,
): GanttTask[] {
  if (deltaDays === 0 || ids.size === 0) return tasks
  return tasks.map((t) => {
    if (!ids.has(t.id)) return t
    if (t.start === null && t.end === null) return t
    return {
      ...t,
      start: t.start === null ? null : addDaysISO(t.start, deltaDays),
      end: t.end === null ? null : addDaysISO(t.end, deltaDays),
    }
  })
}

/** Convenience wrapper: move everything downstream of `sourceId` by N days. */
export function rippleFrom(
  tasks: GanttTask[],
  sourceId: string,
  deltaDays: number,
): GanttTask[] {
  if (deltaDays === 0) return tasks
  return shiftTasksByDays(tasks, collectDownstreamIds(tasks, sourceId), deltaDays)
}

/**
 * True when making `taskId` depend on `candidateId` would close a loop — either
 * directly, or through the parent/child rollup that makes a subtree move as one.
 */
export function wouldCreateCycle(
  tasks: GanttTask[],
  taskId: string,
  candidateId: string,
): boolean {
  if (taskId === candidateId) return true
  const group = relatedGroup(tasks, taskId)
  if (group.has(candidateId)) return true
  // If the candidate already rides along when this task moves, depending on it
  // would make the two chase each other.
  return collectDownstreamIds(tasks, taskId).has(candidateId)
}

/**
 * Tasks that `taskId` could legally be linked to, in document order.
 *
 * Same rule as `wouldCreateCycle`, but the two expensive traversals run once
 * for the row instead of once per candidate — the per-candidate form turns a
 * long task list into an O(n²) render.
 */
export function eligibleDepCandidates(
  tasks: GanttTask[],
  taskId: string,
): GanttTask[] {
  const group = relatedGroup(tasks, taskId)
  const downstream = collectDownstreamIds(tasks, taskId)
  const existing = new Set(tasks.find((t) => t.id === taskId)?.deps ?? [])
  return tasks.filter(
    (t) => !group.has(t.id) && !downstream.has(t.id) && !existing.has(t.id),
  )
}

/** Predecessor ids that still point at a live task, de-duplicated. */
export function validDeps(task: GanttTask, tasks: GanttTask[]): string[] {
  if (!task.deps?.length) return []
  const ids = new Set(tasks.map((t) => t.id))
  return [...new Set(task.deps)].filter((id) => id !== task.id && ids.has(id))
}

/** Drops links pointing at tasks that no longer exist. */
export function pruneDanglingDeps(tasks: GanttTask[]): GanttTask[] {
  const ids = new Set(tasks.map((t) => t.id))
  let changed = false
  const next = tasks.map((t) => {
    if (!t.deps?.length) return t
    const kept = [...new Set(t.deps)].filter((id) => id !== t.id && ids.has(id))
    if (kept.length === t.deps.length) return t
    changed = true
    return { ...t, deps: kept.length ? kept : undefined }
  })
  return changed ? next : tasks
}
