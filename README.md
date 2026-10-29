# PPT TextMerge

**PPT TextMerge** is a lightweight PowerPoint add-in for repairing fragmented text boxes, especially after **PDF → PowerPoint** conversion.

## v0.6 — stable merge engine

The v0.6 engine deliberately avoids recreating text boxes.

Instead, it keeps an existing text box as the merge target, replaces only its text, adjusts its geometry, and deletes the absorbed boxes. It no longer copies fragile PDF-derived margins, horizontal alignment, font/color runs, or other text-frame properties.

### Ribbon commands

- **Fusion paragraphes** — detects logical paragraphs and merges each group into its existing topmost text box.
- **Fusion 1 zone** — merges the whole selection into one existing target box.
- **Fusion colonnes** — detects columns and merges each column into a separate existing text box.
- **Options** — controls paragraph-gap detection, title/body splitting, column detection, compact/preserved spacing, and dehyphenation.

### Important limitation

Office.js does not currently expose PowerPoint's native text-column count. Therefore **Fusion colonnes** creates one stable text box per detected column; it does not create a single native multi-column PowerPoint text frame.

## Production hosting

https://dreadaxe.github.io/ppt-textmerge

No Node/npm/Vite runtime is required for end users.
