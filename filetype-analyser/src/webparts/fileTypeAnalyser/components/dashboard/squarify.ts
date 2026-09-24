export interface IRect {
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface ITreemapInput<T> {
  value: number;
  data: T;
}

export interface ITreemapCell<T> {
  rect: IRect;
  data: T;
}

interface INode<T> {
  area: number;
  data: T;
}

function sumArea<T>(row: INode<T>[]): number {
  let total = 0;
  row.forEach((n) => {
    total += n.area;
  });
  return total;
}

/** Worst aspect ratio in a row laid along a side of the given length (lower is squarer). */
function worst<T>(row: INode<T>[], side: number): number {
  const s = sumArea(row);
  let max = 0;
  let min = Infinity;
  row.forEach((n) => {
    max = Math.max(max, n.area);
    min = Math.min(min, n.area);
  });
  const s2 = s * s;
  const side2 = side * side;
  return Math.max((side2 * max) / s2, s2 / (side2 * min));
}

function layoutRow<T>(row: INode<T>[], rect: IRect, out: ITreemapCell<T>[]): IRect {
  const s = sumArea(row);
  if (rect.w >= rect.h) {
    const colW = rect.h > 0 ? s / rect.h : 0;
    let y = rect.y;
    row.forEach((n) => {
      const h = colW > 0 ? n.area / colW : 0;
      out.push({ rect: { x: rect.x, y, w: colW, h }, data: n.data });
      y += h;
    });
    return { x: rect.x + colW, y: rect.y, w: rect.w - colW, h: rect.h };
  }
  const rowH = rect.w > 0 ? s / rect.w : 0;
  let x = rect.x;
  row.forEach((n) => {
    const w = rowH > 0 ? n.area / rowH : 0;
    out.push({ rect: { x, y: rect.y, w, h: rowH }, data: n.data });
    x += w;
  });
  return { x: rect.x, y: rect.y + rowH, w: rect.w, h: rect.h - rowH };
}

/**
 * Squarified treemap (Bruls, Huizing & van Wijk): lays items out as
 * near-square rectangles whose areas are proportional to their values.
 */
export function squarify<T>(items: ITreemapInput<T>[], rect: IRect): ITreemapCell<T>[] {
  const positive = items.filter((i) => i.value > 0).sort((a, b) => b.value - a.value);
  let total = 0;
  positive.forEach((i) => {
    total += i.value;
  });
  if (total <= 0 || rect.w <= 0 || rect.h <= 0) {
    return [];
  }

  const scale = (rect.w * rect.h) / total;
  const nodes: INode<T>[] = positive.map((i) => ({ area: i.value * scale, data: i.data }));
  const out: ITreemapCell<T>[] = [];
  let remaining = { ...rect };
  let row: INode<T>[] = [];
  let index = 0;

  while (index < nodes.length) {
    const side = Math.min(remaining.w, remaining.h);
    const candidate = row.concat([nodes[index]]);
    if (row.length === 0 || worst(candidate, side) <= worst(row, side)) {
      row = candidate;
      index++;
    } else {
      remaining = layoutRow(row, remaining, out);
      row = [];
    }
  }
  if (row.length > 0) {
    layoutRow(row, remaining, out);
  }
  return out;
}
