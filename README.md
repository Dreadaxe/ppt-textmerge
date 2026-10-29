# PPT TextMerge

**PPT TextMerge** is a lightweight PowerPoint add-in for repairing fragmented text boxes, especially after **PDF → PowerPoint** conversion.

Repository: `Dreadaxe/ppt-textmerge`.

## Core workflow

Select text boxes on a slide and use the PowerPoint ribbon:

- **Fusion paragraphes** — detects columns, visual lines, titles/body transitions and reconstructs logical paragraphs.
- **Fusion 1 zone** — joins the full selection into one text box.
- **Découper** — splits one text box into explicit paragraphs or explicit line breaks.
- **Options** — opens the lightweight settings/diagnostic panel.

## v0.5 highlights

### Smart same-line fragment repair
Adjacent boxes on the same baseline can be reassembled before paragraph reconstruction. Very small horizontal gaps are treated as OCR/PDF fragmentation inside a word; larger gaps become spaces.

### Column-aware paragraph reconstruction
Text boxes are first grouped into visual lines, then columns, then paragraphs. This reduces accidental cross-column merges.

### Reference-box merge
For **Fusion 1 zone**, choose either:
- the union of all selected boxes; or
- the first selected box as the reference geometry.

### Frame preservation
PPT TextMerge can preserve:
- text margins;
- vertical anchoring;
- word wrap;
- auto-fit behavior;
- dominant font, size, bold, italic and color.

### Split text
One multiline text box can be split on explicit paragraph breaks or explicit line breaks.

### Selection diagnostic
The Options pane can report:
`text boxes → visual lines → detected paragraphs · columns · same-line fragments`
without modifying the slide.

### Keyboard shortcuts
On supported PowerPoint versions:
- `Ctrl+Alt+M` — smart paragraph merge
- `Ctrl+Alt+J` — merge into one box
- `Ctrl+Alt+Shift+P` — split into paragraphs
- `Ctrl+Alt+Shift+L` — split into lines

## Production hosting

The add-in is static and hosted at:

https://dreadaxe.github.io/ppt-textmerge

No Node/npm/Vite runtime is required for end users.
