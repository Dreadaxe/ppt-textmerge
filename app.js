(() => {
  "use strict";

  const SETTINGS_KEY = "paramerge.v041.settings";
  const DEFAULTS = {
    paragraphGap: 1.55,
    splitOnStyle: true,
    respectColumns: true,
    layoutMode: "preserve",
    maxBlankLines: 8,
    dehyphenate: true,
    preserveStyles: true,
    preserveParagraphTop: true
  };

  function supports(set, version) {
    try { return Office.context.requirements.isSetSupported(set, version); }
    catch { return false; }
  }

  function loadSettings() {
    try {
      const raw = localStorage.getItem(SETTINGS_KEY);
      return { ...DEFAULTS, ...(raw ? JSON.parse(raw) : {}) };
    } catch {
      return { ...DEFAULTS };
    }
  }

  function saveSettings(settings) {
    localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
  }

  function setStatus(message, kind = "normal") {
    const el = document.getElementById("status");
    if (el) {
      el.textContent = message;
      el.className = `status${kind === "normal" ? "" : " " + kind}`;
    }
  }

  async function showError(message) {
    setStatus(message, "error");
    try {
      if (Office.addin && Office.addin.showAsTaskpane) {
        await Office.addin.showAsTaskpane();
      }
    } catch {}
  }

  const VALID_ALIGNMENTS = new Set([
    "Left", "Center", "Right", "Justify",
    "JustifyLow", "Distributed", "ThaiDistributed"
  ]);

  function normalizeText(s) {
    return String(s || "").replace(/\s+/g, " ").trim();
  }

  function describeOfficeError(error, stage = "") {
    const parts = [];
    if (stage) parts.push(stage);
    if (error && error.code) parts.push(String(error.code));
    if (error && error.message && error.message !== error.code) parts.push(String(error.message));
    const loc = error?.debugInfo?.errorLocation;
    if (loc) parts.push("à " + loc);
    return parts.join(" — ") || String(error);
  }

  function safeTextMargin(value, dimension) {
    if (typeof value !== "number" || !Number.isFinite(value)) return null;
    if (value < 0) return null;

    // Les conversions PDF peuvent exposer des valeurs de marge aberrantes.
    // Une marge supérieure à la moitié de la dimension du cadre n'est pas
    // utile ici et peut être rejetée par PowerPoint lors de la réécriture.
    const limit = Math.max(0, Math.min(72, Number.isFinite(dimension) ? dimension / 2 : 72));
    if (value > limit) return null;

    return value;
  }

  function median(values) {
    if (!values.length) return 0;
    const a = [...values].sort((x, y) => x - y);
    const m = Math.floor(a.length / 2);
    return a.length % 2 ? a[m] : (a[m - 1] + a[m]) / 2;
  }

  function eligibleType(type) {
    return ["TextBox", "GeometricShape", "Placeholder", "Callout"].includes(String(type));
  }

  async function hydrateShapes(context, shapes) {
    const eligible = shapes.filter(s => eligibleType(s.type));

    for (const shape of eligible) {
      shape.textFrame.load("hasText,leftMargin,rightMargin,topMargin,bottomMargin,wordWrap");
      shape.textFrame.textRange.load(
        "text,font/name,font/size,font/bold,font/italic,font/color,paragraphFormat/horizontalAlignment"
      );
    }
    await context.sync();

    const out = [];
    for (const shape of eligible) {
      const tf = shape.textFrame;
      if (!tf.hasText) continue;

      const text = normalizeText(tf.textRange.text);
      if (!text || text.includes("\n") || text.includes("\v")) continue;

      const font = tf.textRange.font;
      out.push({
        id: shape.id,
        text,
        left: shape.left,
        top: shape.top,
        width: shape.width,
        height: shape.height,
        right: shape.left + shape.width,
        bottom: shape.top + shape.height,
        cx: shape.left + shape.width / 2,
        font: {
          name: font.name || null,
          size: font.size || null,
          bold: font.bold == null ? null : font.bold,
          italic: font.italic == null ? null : font.italic,
          color: font.color || null
        },
        align: tf.textRange.paragraphFormat.horizontalAlignment == null
          ? null
          : String(tf.textRange.paragraphFormat.horizontalAlignment),
        frame: {
          leftMargin: safeTextMargin(tf.leftMargin, shape.width),
          rightMargin: safeTextMargin(tf.rightMargin, shape.width),
          topMargin: safeTextMargin(tf.topMargin, shape.height),
          bottomMargin: safeTextMargin(tf.bottomMargin, shape.height)
        }
      });
    }
    return out;
  }

  async function getSelection() {
    return PowerPoint.run(async context => {
      const slides = context.presentation.getSelectedSlides();
      const shapes = context.presentation.getSelectedShapes();

      // Pattern recommandé par Microsoft pour les scoped collections :
      // charger d'abord "items", puis les propriétés de chaque objet.
      slides.load("items");
      shapes.load("items");
      const shapeCount = shapes.getCount();

      try {
        await context.sync();
      } catch (error) {
        throw new Error(describeOfficeError(error, "Lecture de la sélection"));
      }

      if (slides.items.length !== 1) {
        throw new Error("Sélectionne une seule diapositive.");
      }
      if (shapeCount.value < 1 || shapes.items.length < 1) {
        throw new Error("Sélectionne des zones de texte.");
      }

      const slide = slides.items[0];
      slide.load("id");

      for (const shape of shapes.items) {
        shape.load("id,type,left,top,width,height");
      }

      try {
        await context.sync();
      } catch (error) {
        throw new Error(describeOfficeError(error, "Lecture des propriétés des zones"));
      }

      const snapshots = await hydrateShapes(context, shapes.items);
      if (!snapshots.length) {
        throw new Error("Aucune zone de texte exploitable dans la sélection.");
      }

      return { slideId: slide.id, shapes: snapshots };
    });
  }

  function nominalPitch(lines) {
    if (lines.length < 2) {
      return Math.max(8, (lines[0]?.font.size || 10) * 1.15);
    }

    const deltas = lines
      .slice(1)
      .map((x, i) => x.top - lines[i].top)
      .filter(x => x > 1)
      .sort((a, b) => a - b);

    if (!deltas.length) {
      return Math.max(8, median(lines.map(s => s.font.size || 10)) * 1.15);
    }

    const fs = median(lines.map(s => s.font.size || 10));
    const compact = deltas.slice(0, Math.max(1, Math.ceil(deltas.length * 0.60)));
    return Math.max(fs * 0.90, Math.min(median(compact), fs * 1.65));
  }

  function styleBreak(a, b) {
    const sa = a.font.size || 10;
    const sb = b.font.size || 10;
    const size = Math.abs(sa - sb) > Math.max(0.6, Math.max(sa, sb) * 0.06);
    const bold = a.font.bold != null && b.font.bold != null && a.font.bold !== b.font.bold;
    const italic = a.font.italic != null && b.font.italic != null && a.font.italic !== b.font.italic;
    return size || bold || italic;
  }

  function compose(lines, mode, settings) {
    const sorted = [...lines].sort((a, b) => a.top - b.top || a.left - b.left);
    const pitch = nominalPitch(sorted);

    let text = "";
    const spans = [];

    sorted.forEach((line, index) => {
      let sep = "";

      if (index > 0) {
        const prev = sorted[index - 1];
        const dehyphen =
          settings.dehyphenate &&
          /[A-Za-zÀ-ÖØ-öø-ÿ]-$/.test(text) &&
          /^[a-zà-öø-ÿ]/.test(line.text);

        if (dehyphen) {
          text = text.slice(0, -1);
          if (spans.length && spans[spans.length - 1].length > 0) {
            spans[spans.length - 1].length -= 1;
          }
        } else if (mode === "paragraph") {
          sep = " ";
        } else {
          const steps = Math.max(1, Math.round((line.top - prev.top) / Math.max(pitch, 1)));

          if (settings.layoutMode === "preserve") {
            sep = steps <= 1 && !styleBreak(prev, line)
              ? " "
              : "\n".repeat(Math.min(settings.maxBlankLines, Math.max(1, steps)));
          } else {
            sep = (line.top - prev.top) > pitch * 1.55 || styleBreak(prev, line)
              ? "\n"
              : " ";
          }
        }
      }

      text += sep;
      const start = text.length;
      text += line.text;
      spans.push({ start, length: line.text.length, font: line.font });
    });

    return { text, spans, sorted };
  }

  function applyFont(range, font) {
    if (typeof font.name === "string" && font.name && font.name !== "Mixed") {
      range.font.name = font.name;
    }
    if (typeof font.size === "number" && Number.isFinite(font.size) && font.size > 0) {
      range.font.size = font.size;
    }
    if (typeof font.bold === "boolean") range.font.bold = font.bold;
    if (typeof font.italic === "boolean") range.font.italic = font.italic;
    if (typeof font.color === "string" && font.color && font.color !== "Mixed") {
      range.font.color = font.color;
    }
  }

  async function mergeGroup(slideId, ids, mode, selectResult = false) {
    const settings = loadSettings();

    return PowerPoint.run(async context => {
      const slide = context.presentation.slides.getItem(slideId);
      const source = ids.map(id => slide.shapes.getItem(id));

      source.forEach(s => s.load("id,type,left,top,width,height"));
      await context.sync();

      const lines = await hydrateShapes(context, source);
      if (!lines.length) throw new Error("Le groupe ne contient plus de texte exploitable.");

      const { text, spans, sorted } = compose(lines, mode, settings);

      const left = Math.min(...sorted.map(s => s.left));
      const top = Math.min(...sorted.map(s => s.top));
      const right = Math.max(...sorted.map(s => s.right));
      const bottom = Math.max(...sorted.map(s => s.bottom));

      const merged = slide.shapes.addTextBox(text, {
        left,
        top,
        width: Math.max(4, right - left),
        height: Math.max(4, bottom - top)
      });

      merged.name = `PPT TextMerge ${mode}`;
      merged.fill.clear();

      const first = sorted[0];
      if (first.frame.leftMargin != null) merged.textFrame.leftMargin = first.frame.leftMargin;
      if (first.frame.rightMargin != null) merged.textFrame.rightMargin = first.frame.rightMargin;
      if (first.frame.topMargin != null) merged.textFrame.topMargin = first.frame.topMargin;

      // La marge basse est la propriété qui provoque le plus souvent
      // InvalidArgument sur les zones issues d'un PDF. Pour un paragraphe,
      // elle n'est pas nécessaire pour conserver la position visuelle :
      // on privilégie donc la stabilité et on ne la recopie que si elle est sûre.
      if (mode !== "paragraph" && first.frame.bottomMargin != null) {
        merged.textFrame.bottomMargin = first.frame.bottomMargin;
      }

      merged.textFrame.wordWrap = true;

      // PDF-converted line boxes are often vertically centered or auto-sized.
      // Reusing that behavior in a taller paragraph box can move the paragraph.
      // For paragraph fusion, keep the original top coordinate as an invariant.
      if (mode === "paragraph" && settings.preserveParagraphTop) {
        try { merged.textFrame.verticalAlignment = "Top"; } catch {}
        try { merged.textFrame.autoSizeSetting = "AutoSizeNone"; } catch {}
      }

      if (settings.preserveStyles) {
        for (const span of spans) {
          if (span.length <= 0) continue;
          applyFont(
            merged.textFrame.textRange.getSubstring(span.start, span.length),
            span.font
          );
        }
      } else {
        applyFont(merged.textFrame.textRange, first.font);
      }

      if (first.align && VALID_ALIGNMENTS.has(first.align)) {
        merged.textFrame.textRange.paragraphFormat.horizontalAlignment = first.align;
      }

      if (mode === "paragraph" && settings.preserveParagraphTop) {
        // Re-apply the exact source paragraph envelope after text/style changes.
        // The paragraph may extend downward later, but its top must not drift.
        merged.left = left;
        merged.top = top;
        merged.width = Math.max(4, right - left);
        merged.height = Math.max(4, bottom - top);
      }

      merged.load("id");
      try {
        await context.sync();
      } catch (error) {
        throw new Error(describeOfficeError(error, "Création / mise en forme de la zone fusionnée"));
      }

      source.forEach(s => s.delete());
      try {
        await context.sync();
      } catch (error) {
        throw new Error(describeOfficeError(error, "Suppression des anciennes zones"));
      }

      if (selectResult) {
        slide.setSelectedShapes([merged.id]);
        try {
          await context.sync();
        } catch (error) {
          // La fusion elle-même est réussie : un échec de resélection ne doit pas l'annuler.
          console.warn(describeOfficeError(error, "Resélection"));
        }
      }

      return merged.id;
    });
  }

  function clusterColumns(shapes) {
    const settings = loadSettings();
    if (!settings.respectColumns || shapes.length < 4) return [shapes];

    const fs = median(shapes.map(s => s.font.size || 10));
    const tolerance = Math.max(10, fs * 1.55);
    const sorted = [...shapes].sort((a, b) => a.left - b.left || a.top - b.top);
    const clusters = [];

    for (const shape of sorted) {
      let best = -1;
      let bestDistance = Infinity;

      for (let i = 0; i < clusters.length; i++) {
        const anchor = median(clusters[i].map(s => s.left));
        const distance = Math.abs(shape.left - anchor);

        if (distance <= tolerance && distance < bestDistance) {
          best = i;
          bestDistance = distance;
        }
      }

      if (best >= 0) clusters[best].push(shape);
      else clusters.push([shape]);
    }

    const meaningful = clusters.filter(c => c.length >= 2);
    return meaningful.length >= 2 ? meaningful : [shapes];
  }

  function splitParagraphs(lines) {
    const settings = loadSettings();
    const sorted = [...lines].sort((a, b) => a.top - b.top || a.left - b.left);

    if (sorted.length <= 1) return sorted.map(x => [x]);

    const pitch = nominalPitch(sorted);
    const groups = [[sorted[0]]];

    for (let i = 1; i < sorted.length; i++) {
      const prev = sorted[i - 1];
      const cur = sorted[i];

      const gapBreak = (cur.top - prev.top) > pitch * settings.paragraphGap;
      const style = settings.splitOnStyle && styleBreak(prev, cur);

      if (gapBreak || style) groups.push([cur]);
      else groups[groups.length - 1].push(cur);
    }

    return groups;
  }

  async function mergeAllSelection() {
    const data = await getSelection();
    if (data.shapes.length < 2) throw new Error("Sélectionne au moins deux zones de texte.");

    await mergeGroup(
      data.slideId,
      data.shapes.map(s => s.id),
      "all",
      true
    );

    setStatus(`${data.shapes.length} lignes fusionnées en 1 zone.`, "success");
  }

  async function mergeParagraphSelection() {
    const data = await getSelection();
    if (data.shapes.length < 2) throw new Error("Sélectionne au moins deux zones de texte.");

    const columns = clusterColumns(data.shapes);
    const groups = columns.flatMap(col => splitParagraphs(col));

    const mergeable = groups.filter(g => g.length >= 2);
    const untouched = groups
      .filter(g => g.length === 1)
      .flat()
      .map(x => x.id);

    const newIds = [];

    for (const group of mergeable) {
      const id = await mergeGroup(
        data.slideId,
        group.map(x => x.id),
        "paragraph",
        false
      );
      newIds.push(id);
    }

    try {
      await PowerPoint.run(async context => {
        const slide = context.presentation.slides.getItem(data.slideId);
        slide.setSelectedShapes([...untouched, ...newIds]);
        await context.sync();
      });
    } catch (error) {
      console.warn(describeOfficeError(error, "Resélection finale"));
    }

    setStatus(
      `${mergeable.length} paragraphe(s) fusionné(s), ${untouched.length} bloc(s) laissé(s) intact(s).`,
      "success"
    );
  }

  async function runCommand(event, fn) {
    try {
      await fn();
    } catch (e) {
      await showError(
        e instanceof Error && !(e && e.code)
          ? e.message
          : describeOfficeError(e)
      );
    } finally {
      try { event.completed(); } catch {}
    }
  }

  function bindHelpTooltips() {
    const tooltip = document.getElementById("helpTooltip");
    if (!tooltip) return;

    const clamp = (value, min, max) => Math.max(min, Math.min(max, value));

    const show = target => {
      const text = target?.dataset?.tooltip;
      if (!text) return;

      tooltip.textContent = text;
      tooltip.hidden = false;
      tooltip.style.visibility = "hidden";
      tooltip.style.left = "12px";
      tooltip.style.top = "12px";

      const anchor = target.getBoundingClientRect();
      const tip = tooltip.getBoundingClientRect();
      const margin = 8;
      const viewportW = document.documentElement.clientWidth;
      const viewportH = document.documentElement.clientHeight;

      let left = anchor.left + (anchor.width / 2) - (tip.width / 2);
      left = clamp(left, 8, Math.max(8, viewportW - tip.width - 8));

      let top = anchor.top - tip.height - margin;
      if (top < 8) {
        top = anchor.bottom + margin;
      }
      if (top + tip.height > viewportH - 8) {
        top = clamp(anchor.top + (anchor.height / 2) - (tip.height / 2), 8, Math.max(8, viewportH - tip.height - 8));
      }

      tooltip.style.left = `${Math.round(left)}px`;
      tooltip.style.top = `${Math.round(top)}px`;
      tooltip.style.visibility = "visible";
    };

    const hide = () => {
      tooltip.hidden = true;
    };

    document.querySelectorAll(".help").forEach(target => {
      target.addEventListener("mouseenter", () => show(target));
      target.addEventListener("mouseleave", hide);
      target.addEventListener("focus", () => show(target));
      target.addEventListener("blur", hide);
      target.addEventListener("keydown", event => {
        if (event.key === "Escape") hide();
      });
    });

    window.addEventListener("resize", hide);
    document.querySelector("main")?.addEventListener("scroll", hide, { passive: true });
  }

  function bindSettingsUi() {
    const q = id => document.getElementById(id);
    const s = loadSettings();

    const gap = q("paragraphGap");
    if (!gap) return;

    const gapVal = q("paragraphGapValue");
    const split = q("splitOnStyle");
    const cols = q("respectColumns");
    const layout = q("layoutMode");
    const blanks = q("maxBlankLines");
    const blanksVal = q("maxBlankLinesValue");
    const dehy = q("dehyphenate");
    const styles = q("preserveStyles");

    gap.value = String(Math.round(s.paragraphGap * 100));
    gapVal.textContent = `${gap.value} %`;
    split.checked = !!s.splitOnStyle;
    cols.checked = !!s.respectColumns;
    layout.value = s.layoutMode;
    blanks.value = String(s.maxBlankLines);
    blanksVal.textContent = `${blanks.value} lignes`;
    dehy.checked = !!s.dehyphenate;
    styles.checked = !!s.preserveStyles;

    const update = () => {
      saveSettings({
        paragraphGap: Number(gap.value) / 100,
        splitOnStyle: split.checked,
        respectColumns: cols.checked,
        layoutMode: layout.value,
        maxBlankLines: Number(blanks.value),
        dehyphenate: dehy.checked,
        preserveStyles: styles.checked
      });

      gapVal.textContent = `${gap.value} %`;
      blanksVal.textContent = `${blanks.value} lignes`;
      setStatus("Options enregistrées.", "success");
    };

    [gap, split, cols, layout, blanks, dehy, styles]
      .forEach(el => el.addEventListener("input", update));
  }

  Office.onReady(info => {
    try {
      Office.actions.associate(
        "mergeParagraphsCommand",
        event => runCommand(event, mergeParagraphSelection)
      );
      Office.actions.associate(
        "mergeSingleCommand",
        event => runCommand(event, mergeAllSelection)
      );
    } catch (e) {
      console.error(e);
    }

    bindSettingsUi();
    bindHelpTooltips();

    const badge = document.getElementById("apiBadge");
    if (badge) {
      const ok =
        info.host === Office.HostType.PowerPoint &&
        supports("PowerPointApi", "1.5");
      badge.textContent = ok ? "PowerPoint API 1.5+" : "API incompatible";
    }

    setStatus("Prêt.");
  });
})();
