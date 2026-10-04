// Slide IR v2 (EMU) -> editable PPTX.
//
// Writer for the common IR plus pixel_layout_engine.py extensions.
// Keeps PowerPoint-native editable text/table/shapes where practical.
// Unsupported objects are reported instead of silently pretending fidelity.
const fs = require('fs');
const path = require('path');
const pptxgen = require('pptxgenjs');

const EMU_PER_INCH = 914400;
const inch = v => Number(v || 0) / EMU_PER_INCH;

function normalizeColor(value, fallback = '111111') {
  if (!value) return fallback;
  const s = String(value).replace('#', '').toUpperCase();
  return /^[0-9A-F]{6}$/.test(s) ? s : fallback;
}

function firstRun(el) {
  return (el.runs || []).find(r => r && (r.text || r.fontFamily || r.fontSizePt)) || {};
}

function textOptions(el, box) {
  const style = el.style || {};
  const run = firstRun(el);
  const marginPx = Number(style.marginPx || 0);
  // 96dpi: px -> inch, PptxGenJS text margins are points when numeric arrays
  // in several versions; keeping zero is the safest default for deterministic geometry.
  const margin = marginPx === 0 ? 0 : marginPx * 72 / 96;
  return {
    ...box,
    margin,
    fontFace: run.fontFamily || style.fontFamily || 'Yu Gothic',
    fontSize: Number(run.fontSizePt || style.fontSizePt || 18),
    bold: run.bold != null ? !!run.bold : !!style.bold,
    italic: run.italic != null ? !!run.italic : !!style.italic,
    color: normalizeColor(run.color || style.color, '111111'),
    align: style.align || 'left',
    valign: style.valign || 'top',
    breakLine: false,
    fit: 'shrink',
  };
}

function shapeStyle(style = {}) {
  const fillColor = style.fillColor || style.fill || null;
  const lineColor = style.lineColor || style.color || 'D1D5DB';
  return {
    fill: fillColor
      ? { color: normalizeColor(fillColor, 'FFFFFF'), transparency: Number(style.fillTransparency || 0) }
      : { color: 'FFFFFF', transparency: 100 },
    line: {
      color: normalizeColor(lineColor, 'D1D5DB'),
      width: Number(style.widthPt || style.lineWidthPt || 1),
      transparency: Number(style.lineTransparency || 0),
    },
  };
}

async function irToPptx(deck, outputFile) {
  const pptx = new pptxgen();
  const sw = inch(deck.slideSize?.width || 12192000);
  const sh = inch(deck.slideSize?.height || 6858000);
  pptx.defineLayout({ name: 'IR_LAYOUT', width: sw, height: sh });
  pptx.layout = 'IR_LAYOUT';
  pptx.author = deck.meta?.author || 'presentation-toolkit';
  pptx.subject = deck.meta?.subject || '';
  pptx.title = deck.meta?.title || '';

  const unsupported = [];
  const warnings = [];
  const animationReservations = [];

  for (const s of deck.slides || []) {
    const slide = pptx.addSlide();
    if (s.background) {
      slide.background = { color: normalizeColor(s.background, 'FFFFFF') };
    }

    const elements = [...(s.elements || [])].sort((a, b) => (a.z || 0) - (b.z || 0));

    for (const el of elements) {
      const box = {
        x: inch(el.x),
        y: inch(el.y),
        w: inch(el.width),
        h: inch(el.height),
        rotate: Number(el.rotation || 0),
      };

      try {
        if (el.animation) {
          animationReservations.push({
            slide: s.id,
            element: el.id,
            animation: el.animation,
            applied: false,
            backend: 'reserved-for-powerpoint-postprocess',
          });
        }

        if (el.type === 'sp' && (el.text || el.runs?.length)) {
          slide.addText(el.text || '', textOptions(el, box));
          continue;
        }

        if (el.type === 'pic') {
          const src = el.source || {};
          const p = src.path ? path.resolve(src.path) : null;
          if (p && fs.existsSync(p)) {
            slide.addImage({ path: p, ...box });
          } else {
            warnings.push({
              slide: s.id,
              element: el.id,
              type: el.type,
              reason: 'image-path-missing',
              path: p,
            });
          }
          continue;
        }

        if (el.type === 'media') {
          const src = el.source || {};
          const p = src.path ? path.resolve(src.path) : null;
          if (p && fs.existsSync(p)) {
            slide.addMedia({
              type: src.mediaType === 'audio' ? 'audio' : 'video',
              path: p,
              x: box.x,
              y: box.y,
              w: box.w,
              h: box.h,
            });
          } else {
            warnings.push({
              slide: s.id,
              element: el.id,
              type: el.type,
              reason: 'media-path-missing',
              path: p,
            });
          }
          continue;
        }

        if (el.table && el.table.rows && el.table.rows.length) {
          const rows = el.table.rows.map((r, ridx) =>
            r.cells.map(c => {
              const run = (c.runs || []).find(x => x && (x.text || x.fontFamily || x.fontSizePt)) || {};
              return {
                text: c.text || '',
                options: {
                  fontFace: run.fontFamily || el.style?.fontFamily || 'Yu Gothic',
                  fontSize: Number(run.fontSizePt || el.style?.fontSizePt || 14),
                  bold: run.bold != null ? !!run.bold : ridx === 0,
                  italic: !!run.italic,
                  color: normalizeColor(run.color || el.style?.color, '111111'),
                  colspan: c.gridSpan > 1 ? c.gridSpan : undefined,
                  rowspan: c.rowSpan > 1 ? c.rowSpan : undefined,
                  align: el.style?.align || 'center',
                  valign: el.style?.valign || 'middle',
                },
              };
            })
          );

          const colW =
            el.table.colWidths && el.table.colWidths.length > 0
              ? el.table.colWidths.map(w => inch(w))
              : undefined;

          slide.addTable(rows, {
            x: box.x,
            y: box.y,
            w: box.w,
            h: box.h,
            colW,
            border: {
              type: 'solid',
              pt: Number(el.style?.borderWidthPt || 1),
              color: normalizeColor(el.style?.borderColor, 'CBD5E1'),
            },
            fill: { color: normalizeColor(el.style?.fillColor, 'FFFFFF') },
            margin: Number(el.style?.cellMarginPt || 3),
            autoFit: false,
          });
          continue;
        }

        if (el.type === 'line') {
          const style = el.style || {};
          slide.addShape(pptx.ShapeType.line, {
            ...box,
            line: {
              color: normalizeColor(style.color || style.lineColor, '111111'),
              width: Number(style.widthPt || 1),
              transparency: Number(style.lineTransparency || 0),
              beginArrowType: style.beginArrowType || 'none',
              endArrowType: style.endArrowType || 'none',
            },
          });
          continue;
        }

        if (el.type === 'box' || el.type === 'rect' || el.type === 'shape') {
          slide.addShape(pptx.ShapeType.rect, {
            ...box,
            ...shapeStyle(el.style || {}),
          });
          continue;
        }

        unsupported.push({
          slide: s.id,
          element: el.id,
          name: el.name,
          type: el.type,
          reason: 'not-yet-lossless',
        });
      } catch (error) {
        unsupported.push({
          slide: s.id,
          element: el.id,
          name: el.name,
          type: el.type,
          reason: 'writer-error',
          error: String(error && error.message ? error.message : error),
        });
      }
    }
  }

  fs.mkdirSync(path.dirname(path.resolve(outputFile)), { recursive: true });
  await pptx.writeFile({ fileName: path.resolve(outputFile) });
  return {
    output: path.resolve(outputFile),
    unsupported,
    warnings,
    animationReservations,
  };
}

async function main() {
  const [input, output, report] = process.argv.slice(2);
  if (!input || !output) {
    throw new Error('Usage: node src/ir_to_pptx.js deck.json output.pptx [report.json]');
  }
  const result = await irToPptx(JSON.parse(fs.readFileSync(input, 'utf8')), output);
  if (report) {
    fs.mkdirSync(path.dirname(path.resolve(report)), { recursive: true });
    fs.writeFileSync(report, JSON.stringify(result, null, 2));
  }
  console.log(JSON.stringify(result, null, 2));
}

if (require.main === module) {
  main().catch(e => {
    console.error(e);
    process.exitCode = 1;
  });
}

module.exports = { irToPptx, EMU_PER_INCH };
