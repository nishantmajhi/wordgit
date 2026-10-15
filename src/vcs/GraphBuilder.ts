/**
 * GraphBuilder.ts - commit DAG layout for the WordGit branch graph.
 *
 * Given the repository's commits and refs, works out an x/y position for every
 * commit and the edges between them, ready to render as SVG. There is no SVG
 * here and no DOM: this decides where things go, the task pane draws them.
 *
 * Layout:
 *   1. Order the commits newest-first, and never place a commit below its own
 *      parent (see `topoOrder`).
 *   2. row = position in that order, 0 = newest.
 *   3. Each branch tip claims a lane; a commit takes the lane of the child
 *      that reached it first.
 *   4. Emit nodes and edges, following both parents of a merge.
 */

import { ROOT_PARENT } from "./RepoFormat";

export interface GraphCommit {
  hash: string;
  parent: string;
  /** Second parent for a merge commit, otherwise "". */
  parent2: string;
  message: string;
  timestamp: string;
  author: string;
}

export interface GraphRef {
  name: string;
  tip: string;
}

export interface GraphNode {
  hash: string;
  message: string;
  ts: string;
  author: string;
  col: number; // 0-indexed lane
  row: number; // 0 = newest
  refs: string[]; // branch names pointing here
  isRoot: boolean;
  isMerge: boolean;
}

export interface GraphEdge {
  fromHash: string;
  toHash: string;
  fromCol: number;
  fromRow: number;
  toCol: number;
  toRow: number;
}

export interface GraphLayout {
  nodes: GraphNode[];
  edges: GraphEdge[];
  cols: number;
  rows: number;
}

function parentsOf(c: GraphCommit, byHash: Map<string, GraphCommit>): string[] {
  return [c.parent, c.parent2].filter((h) => h && h !== ROOT_PARENT && byHash.has(h));
}

function timeOf(c: GraphCommit): number {
  const t = Date.parse(c.timestamp);
  return Number.isNaN(t) ? 0 : t;
}

/**
 * Order the reachable commits newest-first, emitting a commit only once every
 * commit that has it as a parent has already been emitted.
 *
 * The original used a breadth-first queue from the branch tips. That happened
 * to keep parents below children on a linear history, but it interleaved
 * branches by depth rather than by date, so two branches of different lengths
 * came out shuffled together in an order that matched neither the graph nor the
 * clock. Choosing the newest ready commit at each step gives date order where
 * the graph allows it and graph order where it does not.
 */
function topoOrder(commits: GraphCommit[], tips: string[]): string[] {
  const byHash = new Map<string, GraphCommit>();
  for (const c of commits) byHash.set(c.hash, c);

  // Restrict to what is reachable from the tips.
  const reachable = new Set<string>();
  const stack = tips.filter((h) => byHash.has(h));
  while (stack.length > 0) {
    const h = stack.pop()!;
    if (reachable.has(h)) continue;
    reachable.add(h);
    for (const p of parentsOf(byHash.get(h)!, byHash)) {
      if (!reachable.has(p)) stack.push(p);
    }
  }

  const pending = new Map<string, number>();
  for (const h of reachable) pending.set(h, 0);
  for (const h of reachable) {
    for (const p of parentsOf(byHash.get(h)!, byHash)) {
      if (pending.has(p)) pending.set(p, pending.get(p)! + 1);
    }
  }

  const ready = new Set<string>();
  for (const [h, n] of pending) if (n === 0) ready.add(h);

  const order: string[] = [];
  while (ready.size > 0) {
    let pick: GraphCommit | null = null;
    for (const h of ready) {
      const c = byHash.get(h)!;
      if (!pick || timeOf(c) > timeOf(pick) || (timeOf(c) === timeOf(pick) && c.hash > pick.hash)) {
        pick = c;
      }
    }
    ready.delete(pick!.hash);
    order.push(pick!.hash);
    for (const p of parentsOf(pick!, byHash)) {
      if (!pending.has(p)) continue;
      const left = pending.get(p)! - 1;
      pending.set(p, left);
      if (left === 0) ready.add(p);
    }
  }
  return order;
}

export function buildGraph(commits: GraphCommit[], refs: GraphRef[]): GraphLayout {
  if (commits.length === 0) return { nodes: [], edges: [], cols: 1, rows: 0 };

  const byHash = new Map<string, GraphCommit>();
  for (const c of commits) byHash.set(c.hash, c);

  const order = topoOrder(
    commits,
    refs.map((r) => r.tip)
  );
  if (order.length === 0) return { nodes: [], edges: [], cols: 1, rows: 0 };

  // Lane assignment: each ref tip claims a lane, then a commit takes the lane
  // of whichever child reached it first in the order above.
  const laneNames: string[] = [];
  const commitLane = new Map<string, number>();
  for (const ref of refs) {
    if (!byHash.has(ref.tip) || commitLane.has(ref.tip)) continue;
    commitLane.set(ref.tip, laneNames.length);
    laneNames.push(ref.name);
  }

  for (const hash of order) {
    if (!commitLane.has(hash)) commitLane.set(hash, 0);
    for (const p of parentsOf(byHash.get(hash)!, byHash)) {
      if (!commitLane.has(p)) commitLane.set(p, commitLane.get(hash)!);
    }
  }

  const hashRefs = new Map<string, string[]>();
  for (const ref of refs) {
    if (!byHash.has(ref.tip)) continue;
    const list = hashRefs.get(ref.tip) ?? [];
    list.push(ref.name);
    hashRefs.set(ref.tip, list);
  }

  const nodes: GraphNode[] = order.map((hash, row) => {
    const c = byHash.get(hash)!;
    return {
      hash,
      message: c.message,
      ts: c.timestamp,
      author: c.author,
      col: commitLane.get(hash) ?? 0,
      row,
      refs: hashRefs.get(hash) ?? [],
      isRoot: c.parent === ROOT_PARENT,
      isMerge: parentsOf(c, byHash).length > 1,
    };
  });

  const nodeByHash = new Map<string, GraphNode>();
  for (const n of nodes) nodeByHash.set(n.hash, n);

  const edges: GraphEdge[] = [];
  for (const hash of order) {
    const from = nodeByHash.get(hash)!;
    for (const parent of parentsOf(byHash.get(hash)!, byHash)) {
      const to = nodeByHash.get(parent);
      if (!to) continue;
      edges.push({
        fromHash: hash,
        toHash: parent,
        fromCol: from.col,
        fromRow: from.row,
        toCol: to.col,
        toRow: to.row,
      });
    }
  }

  return { nodes, edges, cols: laneNames.length || 1, rows: nodes.length };
}
