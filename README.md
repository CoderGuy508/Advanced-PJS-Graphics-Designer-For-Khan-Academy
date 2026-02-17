# PJS Graphics Designer

A browser-based ProcessingJS graphics editor for building art visually and exporting compact Khan Academy-compatible `drawArt(...)` code.

## Features

- Editable canvas size (`100` to `2000` px for width/height)
- Tools: `select`, `brush`, `pencil`, `glow`, `spray`, `eraser`, `line`, `rect`, `ellipse`, `arc`, `triangle`, `quad`, `bezier`, `vertex`, `curve`, `fill`
- Separate controls for:
  - Fill color
  - Stroke color
  - Stroke weight
  - Brush size
- Arc controls (`start`/`end` angle)
- Optional shape gradients (`left-right`, `right-left`, `top-bottom`, `bottom-top`, `center-out`, `center-in`)
- Triangle gradients now export with a dedicated HQ helper for Khan PJS.
- Gradient export is strongest for `rect`, `ellipse`, and `triangle`; complex polygons/paths export as solid fills.
- Path tools:
  - `vertex` (click points, `Enter` to finish, `Esc` to cancel)
  - `curve` (curve-style vertex path for PJS `curveVertex`)
- Selection editing (single + multi):
  - Click to select, drag-box marquee to select multiple objects
  - Group move/resize/rotate with a shared selection box + grips
  - Drag control points on supported single selections (`line`, `arc`, `triangle`, `quad`, `bezier`, `vertex`, `curve`)
  - Apply current style controls to all selected objects
  - Delete selected objects (button or Delete/Backspace)
  - Copy/paste selected object with `Ctrl/Cmd+C` and `Ctrl/Cmd+V`
  - Layer ordering: forward one, back one, to front, to back
- Color picker, brush size, opacity slider, and shape fill toggle
- Undo/redo
- Autosave + restore using `localStorage`
- Compact PJS export in a reusable function: `drawArt(x, y, w, h)` (radian-safe output)
- Optional image-mode export for pixel-accurate results
- Download artwork as PNG

## Run Locally

```bash
cd "/Users/<your username here>/Documents/untitled folder/pjs-graphics-app"
python3 -m http.server 8787
```

Open [http://127.0.0.1:8787](http://127.0.0.1:8787).

## GitHub Pages

Use the contents of this folder as your site root, or set this folder as the publish directory in your Pages workflow.

## Notes

- Fill-bucket edits are saved/restored in the app but skipped in exported PJS output.
- Exported function uses normalized drawing data so the same art can be drawn at any size.
- For multi-selection, gradient controls are disabled if any selected object type cannot use gradients (for example `arc` or `vertexPath`).
- Created entirely by Codex by OpenAI
