import { describe, expect, it } from 'vitest';
import type { Stroke, TableElement, TextBody } from './types.js';
import { applyStroke, renderTable, renderTextBody, resolveShapeFill } from './renderer.js';

function recordingContext() {
  const linear: number[][] = [];
  const gradient = { addColorStop() {} } as unknown as CanvasGradient;
  const ctx = {
    canvas: {},
    strokeStyle: '',
    lineWidth: 0,
    lineCap: 'butt',
    lineJoin: 'miter',
    miterLimit: 10,
    setLineDash() {},
    createLinearGradient(...args: number[]) {
      linear.push(args);
      return gradient;
    },
    createRadialGradient() {
      return gradient;
    },
  } as unknown as CanvasRenderingContext2D;
  return { ctx, gradient, linear };
}

describe('PPTX DrawingML gradient projection', () => {
  it('counter-rotates rotWithShape=false fills using the authored shape rotation', () => {
    const { ctx, linear } = recordingContext();
    resolveShapeFill({
      fillType: 'gradient',
      stops: [{ position: 0, color: '000000' }, { position: 1, color: 'FFFFFF' }],
      angle: 90,
      gradType: 'linear',
      rotWithShape: false,
    }, ctx, 0, 0, 100, 100, 90);

    [0, 50, 100, 50].forEach((value, index) => {
      expect(linear[0][index]).toBeCloseTo(value);
    });
  });

  it('uses the complete gradient line paint instead of its flat fallback color', () => {
    const { ctx, gradient, linear } = recordingContext();
    const stroke: Stroke = {
      color: 'FF0000',
      width: 12_700,
      fill: {
        fillType: 'gradient',
        stops: [{ position: 0, color: '000000' }, { position: 1, color: 'FFFFFF' }],
        angle: 0,
        gradType: 'linear',
      },
    };

    applyStroke(ctx, stroke, 1 / 12_700, { x: 10, y: 20, w: 100, h: 50 });

    expect(linear).toHaveLength(1);
    expect(ctx.strokeStyle).toBe(gradient);
  });

  it('paints a gradient text run across font slots and keeps the adjacent solid run separate', () => {
    const { ctx, gradient, linear } = recordingContext();
    const paints: Array<[string, typeof ctx.fillStyle]> = [];
    Object.assign(ctx, {
      save() {}, restore() {}, translate() {}, beginPath() {}, rect() {}, clip() {},
      measureText: (text: string) => ({ width: text.length * 10, actualBoundingBoxAscent: 8 }),
      fillText: (text: string) => paints.push([text, ctx.fillStyle]),
    });
    const run = {
      type: 'text', bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 20, fontFamily: 'Arial', fontFamilyEa: 'SimSun',
    } as const;
    const body = {
      paragraphs: [{ alignment: 'l', marL: 0, marR: 0, indent: 0, lvl: 0, bullet: { type: 'none' }, runs: [
        { ...run, text: 'AB中', glyphFill: { fillType: 'gradient', gradType: 'linear', angle: 0,
          stops: [{ position: 0, color: 'FF0000' }, { position: 1, color: '0000FF' }] } },
        { ...run, text: 'C', color: '008000' },
      ] }],
      defaultFontSize: 20, lIns: 0, rIns: 0, tIns: 0, bIns: 0, verticalAnchor: 't',
    } as TextBody;
    renderTextBody(ctx, body, 0, 0, 200, 100, 1 / 12700, null);
    expect(paints).toEqual([['AB', gradient], ['中', gradient], ['C', 'rgba(0,128,0,1)']]);
    expect(linear).toHaveLength(1);
    expect(linear[0][2] - linear[0][0]).toBeCloseTo(30);
  });

  it('uses the complete gradient paint for table cells', () => {
    const { ctx, gradient, linear } = recordingContext();
    let painted: string | CanvasGradient | CanvasPattern = '';
    Object.assign(ctx, {
      save() {},
      restore() {},
      translate() {},
      rotate() {},
      scale() {},
      fillRect() { painted = ctx.fillStyle; },
    });
    const table: TableElement = {
      type: 'table',
      x: 0,
      y: 0,
      width: 100,
      height: 50,
      rotation: 0,
      flipH: false,
      flipV: false,
      cols: [100],
      rows: [{
        height: 50,
        cells: [{
          textBody: null,
          fill: {
            fillType: 'gradient',
            stops: [{ position: 0, color: '000000' }, { position: 1, color: 'FFFFFF' }],
            angle: 0,
            gradType: 'linear',
          },
          borderL: null,
          borderR: null,
          borderT: null,
          borderB: null,
          gridSpan: 1,
          rowSpan: 1,
          hMerge: false,
          vMerge: false,
        }],
      }],
    };

    renderTable(ctx, table, 1);

    expect(linear).toHaveLength(1);
    expect(painted).toBe(gradient);
  });
});
