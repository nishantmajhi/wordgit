/**
 * graphSvg.ts - draws the commit DAG that `GraphBuilder` laid out.
 *
 * Pure string building, no DOM. `GraphBuilder` decides where each node goes;
 * this decides what it looks like.
 *
 * Layout model - each node is a horizontal row, all elements centred on Y:
 *
 *   [pills right-aligned to circle] -- [circle] -- [hash / message / timestamp]
 *
 *   Pills:   right edge = x - PILL_CIRCLE_GAP, vertically centred on y
 *   Circle:  centre at (x, y)
 *   Text:    left edge = x + R + TEXT_LEFT_GAP, three lines centred on y:
 *              hash      at y - 9
 *              message   at y + 2
 *              timestamp at y + 13
 *
 * Lane width is computed from the widest pill across all branch names, so text
 * from one lane can never bleed into the next lane's circle.
 */

import { GraphLayout } from "../vcs/GraphBuilder";
import { escapeHtml } from "./render";

const LANE_COLOURS = ["#0078d4", "#107c10", "#8764b8", "#d04a02", "#00b7c3", "#c4262e"];
const R = 6; // circle radius
const PAD_L = 12; // left edge margin
const PAD_T = 16; // top/bottom padding
const ROW_H = 36; // vertical distance between node centres
const PILL_H = 14; // pill height
const PILL_CIRCLE_GAP = 6; // gap between pill right edge and circle left edge
const TEXT_LEFT_GAP = 8; // gap between circle right edge and first text char
const LANE_SAFETY_PAD = 40; // extra px added to lane width beyond the pill

/** Approximate SVG text width at 9px Segoe UI - used for pill sizing. */
export function pillWidth(name: string): number {
  // 5.5px average char width at 9px, 10px horizontal padding total
  return Math.max(Math.ceil(name.length * 5.5) + 10, 20);
}

function colour(col: number): string {
  return LANE_COLOURS[col % LANE_COLOURS.length];
}

/** Commit timestamp for the bottom line of a node, or "" if unparseable. */
export function graphTimestamp(iso: string): string {
  if (!iso) return "";
  const d = new Date(iso);
  if (isNaN(d.getTime())) return "";
  const p = (n: number) => n.toString().padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(
    d.getMinutes()
  )}`;
}

export function truncate(s: string, max: number): string {
  return s.length > max ? s.slice(0, max - 1) + "…" : s;
}

export interface GraphTheme {
  dark: boolean;
}

export function renderGraphSvg(layout: GraphLayout, theme: GraphTheme): string {
  // Lane width from the widest pill in each lane, plus the cumulative offsets.
  const colPillWidths: number[] = new Array(layout.cols).fill(0);
  for (const node of layout.nodes) {
    for (const name of node.refs) {
      const pw = pillWidth(name);
      if (pw > colPillWidths[node.col]) colPillWidths[node.col] = pw;
    }
  }

  const colW = colPillWidths.map((pw) =>
    Math.max(pw + PILL_CIRCLE_GAP + R + LANE_SAFETY_PAD, 60)
  );

  const colCircleX: number[] = [];
  let xCursor = PAD_L;
  for (let col = 0; col < layout.cols; col++) {
    colCircleX.push(xCursor + colPillWidths[col] + PILL_CIRCLE_GAP + R);
    xCursor += colW[col];
  }

  const gcX = (col: number) => colCircleX[col] ?? PAD_L + R;
  const gcY = (row: number) => PAD_T + row * ROW_H;

  const rightmostX = colCircleX[layout.cols - 1] ?? PAD_L + R;
  const W = Math.max(rightmostX + R + TEXT_LEFT_GAP + 150, 220);
  const H = PAD_T + layout.rows * ROW_H + PAD_T;

  const textFill = theme.dark ? "#d0d0d0" : "#1a1a1a";
  const mutedFill = theme.dark ? "#707070" : "#6b6b6b";
  const bgFill = theme.dark ? "#252525" : "#ffffff";
  const pillBg = theme.dark ? "rgba(0,120,212,0.22)" : "rgba(0,120,212,0.10)";

  const lines: string[] = [
    `<svg xmlns="http://www.w3.org/2000/svg" width="100%" viewBox="0 0 ${W} ${H}" style="display:block;overflow:visible">`,
    `<rect width="${W}" height="${H}" fill="${bgFill}" rx="6"/>`,
  ];

  for (const e of layout.edges) {
    const x1 = gcX(e.fromCol);
    const y1 = gcY(e.fromRow);
    const x2 = gcX(e.toCol);
    const y2 = gcY(e.toRow);
    const stroke = colour(e.toCol);
    if (e.fromCol === e.toCol) {
      lines.push(
        `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" ` +
          `stroke="${stroke}" stroke-width="2" opacity="0.7"/>`
      );
    } else {
      const midY = (y1 + y2) / 2;
      lines.push(
        `<path d="M ${x1} ${y1} C ${x1} ${midY}, ${x2} ${midY}, ${x2} ${y2}" ` +
          `fill="none" stroke="${stroke}" stroke-width="2" opacity="0.65"/>`
      );
    }
  }

  for (const n of layout.nodes) {
    const x = gcX(n.col);
    const y = gcY(n.row);
    const laneColour = colour(n.col);
    const isTip = n.refs.length > 0;

    lines.push(
      `<circle cx="${x}" cy="${y}" r="${R}" ` +
        `fill="${isTip ? laneColour : bgFill}" stroke="${laneColour}" stroke-width="2"/>`
    );

    // Pills, left of the circle and right-aligned to it, stacked right to left.
    let pillRightEdge = x - PILL_CIRCLE_GAP;
    for (let i = n.refs.length - 1; i >= 0; i--) {
      const refName = n.refs[i];
      const pw = pillWidth(refName);
      const px = pillRightEdge - pw;
      const py = y - PILL_H / 2;
      lines.push(
        `<rect x="${px}" y="${py}" width="${pw}" height="${PILL_H}" rx="7" ` +
          `fill="${pillBg}" stroke="rgba(0,120,212,0.35)" stroke-width="0.8"/>`
      );
      lines.push(
        `<text x="${px + pw / 2}" y="${py + PILL_H - 3}" text-anchor="middle" ` +
          `font-family="Segoe UI,system-ui,sans-serif" font-size="9" font-weight="600" fill="${laneColour}">` +
          `${escapeHtml(refName)}</text>`
      );
      pillRightEdge = px - 4;
    }

    const tx = x + R + TEXT_LEFT_GAP;
    lines.push(
      `<text x="${tx}" y="${y - 9}" ` +
        `font-family="Cascadia Code,Consolas,monospace" font-size="9" fill="${mutedFill}">` +
        `${escapeHtml(n.hash.slice(0, 6))}</text>`
    );
    lines.push(
      `<text x="${tx}" y="${y + 2}" ` +
        `font-family="Segoe UI,system-ui,sans-serif" font-size="10" ` +
        `fill="${textFill}" font-weight="${isTip ? 600 : 400}">` +
        `${escapeHtml(truncate(n.message || "(no message)", 22))}</text>`
    );
    lines.push(
      `<text x="${tx}" y="${y + 13}" ` +
        `font-family="Segoe UI,system-ui,sans-serif" font-size="9" fill="${mutedFill}">` +
        `${escapeHtml(graphTimestamp(n.ts))}</text>`
    );
  }

  lines.push("</svg>");
  return lines.join("\n");
}
