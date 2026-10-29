(() => {
  "use strict";

  const SETTINGS_KEY = "ppt-textmerge.v060.settings";
  const LEGACY_SETTINGS_KEY = "paramerge.v041.settings";

  const DEFAULTS = {
    paragraphGap: 1.55,
    splitOnStyle: true,
    respectColumns: true,
    layoutMode: "compact",
    maxBlankLines: 8,
    dehyphenate: true
  };

  function supports(set, version) {
    try {
      return Office.context.requirements.isSetSupported(set, version);
    } catch {
      return false;
    }
  }

  function loadSettings() {
    try {
      const raw =
        localStorage.getItem(SETTINGS_KEY) ||
        localStorage.getItem(LEGACY_SETTINGS_KEY);
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
    if (!el) return;

    el.textContent = message;
    el.className = `status${kind === "normal" ? "" : " " + kind}`;
  }

  async function showError(message) {
    setStatus(message, "error");
    try {
      if (Office.addin?.showAsTaskpane) {
        await Office.addin.showAsTaskpane();
      }
    } catch {}
  }

  function describeOfficeError(error, stage = "") {
    const parts = [];
    if (stage) parts.push(stage);
    if (error?.code) parts.push(String(error.code));
    if (error?.message && error.message !== error.code) {
      parts.push(String(error.message));
    }
    if (error?.debugInfo?.errorLocation) {
      parts.push("à " + error.debugInfo.errorLocation);
    }
    return parts.join(" — ") || String(error);
  }

  function normalizeText(value) {
    return String(value || "").replace(/\s+/g, " ").trim();
  }

  function median(values) {
    if (!values.length) return 0;
    const sorted = [...values].sort((a, b) => a - b);
    const middle = Math.floor(sorted.length / 2);
    return sorted.length % 2
      ? sorted[middle]
      : (sorted[middle - 1] + sorted[middle]) / 2;
  }

  function eligibleType(type) {
    return ["TextBox", "GeometricShape", "Placeholder", "Callout"].includes(
      String(type)
    );
  }

  async function hydrateShapes(context, shapes) {
    const eligible = shapes.filter(shape => eligibleType(shape.type));

    for (const shape of eligible) {
      shape.textFrame.load("hasText");
      shape.textFrame.textRange.load(
        "text,font/size,font/bold,font/italic"
      );
    }

    try {
      await context.sync();
    } catch (error) {
      throw new Error(describeOfficeError(error, "Lecture du texte"));
    }

    const snapshots = [];

    for (const shape of eligible) {
      if (!shape.textFrame.hasText) continue;

      const rawText = String(shape.textFrame.textRange.text || "").trim();
      const text = normalizeText(rawText);
      if (!text) continue;

      const font = shape.textFrame.textRange.font;
      snapshots.push({
        id: shape.id,
        text,
        rawText,
        lineCount: Math.max(
          1,
          rawText.split(/\r\n|\r|\n|\v/).filter(Boolean).length
        ),
        left: shape.left,
        top: shape.top,
        width: shape.width,
        height: shape.height,
        right: shape.left + shape.width,
        bottom: shape.top + shape.height,
        font: {
          size:
            typeof font.size === "number" && Number.isFinite(font.size)
              ? font.size
              : null,
          bold: typeof font.bold === "boolean" ? font.bold : null,
          italic: typeof font.italic === "boolean" ? font.italic : null
        }
      });
    }

    return snapshots;
  }

  async function getSelection() {
    return PowerPoint.run(async context => {
      const slides = context.presentation.getSelectedSlides();
      const shapes = context.presentation.getSelectedShapes();

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
        throw new Error(
          describeOfficeError(error, "Lecture des propriétés des zones")
        );
      }

      const snapshots = await hydrateShapes(context, shapes.items);

      if (!snapshots.length) {
        throw new Error("Aucune zone de texte exploitable dans la sélection.");
      }

      return {
        slideId: slide.id,
        shapes: snapshots
      };
    });
  }

  function nominalPitch(lines) {
    if (lines.length < 2) {
      return Math.max(8, (lines[0]?.font?.size || 10) * 1.15);
    }

    const ordered = [...lines].sort(
      (a, b) => a.top - b.top || a.left - b.left
    );

    const deltas = ordered
      .slice(1)
      .map((line, index) => line.top - ordered[index].top)
      .filter(value => value > 1)
      .sort((a, b) => a - b);

    const fontSize = median(lines.map(line => line.font.size || 10));

    if (!deltas.length) {
      return Math.max(8, fontSize * 1.15);
    }

    const compact = deltas.slice(
      0,
      Math.max(1, Math.ceil(deltas.length * 0.6))
    );

    return Math.max(
      fontSize * 0.9,
      Math.min(median(compact), fontSize * 1.65)
    );
  }

  function effectiveLineHeight(line, pitch) {
    const fontSize = line?.font?.size || 10;
    const logicalLines = Math.max(1, line?.lineCount || 1);
    const cap = Math.max(
      fontSize * 1.8 * logicalLines,
      pitch * 1.3 * logicalLines,
      8
    );

    if (
      typeof line.height === "number" &&
      Number.isFinite(line.height) &&
      line.height > 0
    ) {
      return Math.min(line.height, cap);
    }

    return cap;
  }

  function styleBreak(a, b) {
    const sa = a.font.size || 10;
    const sb = b.font.size || 10;

    const sizeBreak =
      Math.abs(sa - sb) > Math.max(0.6, Math.max(sa, sb) * 0.06);
    const boldBreak =
      a.font.bold != null &&
      b.font.bold != null &&
      a.font.bold !== b.font.bold;
    const italicBreak =
      a.font.italic != null &&
      b.font.italic != null &&
      a.font.italic !== b.font.italic;

    return sizeBreak || boldBreak || italicBreak;
  }

  function compose(lines, mode, settings) {
    const ordered = [...lines].sort(
      (a, b) => a.top - b.top || a.left - b.left
    );
    const pitch = nominalPitch(ordered);

    let text = "";

    ordered.forEach((line, index) => {
      if (index === 0) {
        text = line.text;
        return;
      }

      const previous = ordered[index - 1];

      const dehyphenate =
        settings.dehyphenate &&
        /[A-Za-zÀ-ÖØ-öø-ÿ]-$/.test(text) &&
        /^[a-zà-öø-ÿ]/.test(line.text);

      if (dehyphenate) {
        text = text.slice(0, -1) + line.text;
        return;
      }

      if (mode === "paragraph") {
        text += " " + line.text;
        return;
      }

      if (mode === "column") {
        const paragraphBreak =
          line.top - previous.top > pitch * settings.paragraphGap ||
          (settings.splitOnStyle && styleBreak(previous, line));

        text += (paragraphBreak ? "\n\n" : " ") + line.text;
        return;
      }

      const verticalSteps = Math.max(
        1,
        Math.round((line.top - previous.top) / Math.max(pitch, 1))
      );

      if (settings.layoutMode === "preserve") {
        const separator =
          verticalSteps <= 1 && !styleBreak(previous, line)
            ? " "
            : "\n".repeat(
                Math.min(
                  settings.maxBlankLines,
                  Math.max(1, verticalSteps)
                )
              );
        text += separator + line.text;
      } else {
        const paragraphBreak =
          line.top - previous.top > pitch * settings.paragraphGap ||
          styleBreak(previous, line);
        text += (paragraphBreak ? "\n" : " ") + line.text;
      }
    });

    return {
      text,
      ordered,
      pitch
    };
  }

  function groupEnvelope(lines, pitch) {
    const left = Math.min(...lines.map(line => line.left));
    const top = Math.min(...lines.map(line => line.top));
    const right = Math.max(...lines.map(line => line.right));
    const visualBottom = Math.max(
      ...lines.map(
        line => line.top + effectiveLineHeight(line, pitch)
      )
    );

    return {
      left,
      top,
      width: Math.max(4, right - left),
      height: Math.max(4, visualBottom - top)
    };
  }

  async function mergeGroupStable(
    slideId,
    ids,
    mode,
    selectResult = false
  ) {
    const settings = loadSettings();

    return PowerPoint.run(async context => {
      const slide = context.presentation.slides.getItem(slideId);
      const source = ids.map(id => slide.shapes.getItem(id));

      for (const shape of source) {
        shape.load("id,type,left,top,width,height");
      }

      try {
        await context.sync();
      } catch (error) {
        throw new Error(
          describeOfficeError(error, "Lecture du groupe à fusionner")
        );
      }

      const lines = await hydrateShapes(context, source);

      if (!lines.length) {
        throw new Error("Le groupe ne contient plus de texte exploitable.");
      }

      const { text, ordered, pitch } = compose(lines, mode, settings);
      const targetSnapshot = ordered[0];
      const target = slide.shapes.getItem(targetSnapshot.id);
      const envelope = groupEnvelope(ordered, pitch);

      // Principe v0.6 : on conserve la zone existante la plus haute.
      // Aucun style, marge ou alignement horizontal n'est recopié.
      target.textFrame.textRange.text = text;
      target.left = envelope.left;
      target.top = envelope.top;
      target.width = envelope.width;
      target.height = envelope.height;

      // Ces trois propriétés sont nécessaires pour conserver une géométrie
      // prévisible, sans recopier les réglages fragiles des zones PDF.
      target.textFrame.wordWrap = true;
      target.textFrame.autoSizeSetting = "AutoSizeNone";
      target.textFrame.verticalAlignment = "Top";

      try {
        await context.sync();
      } catch (error) {
        throw new Error(
          describeOfficeError(error, "Mise à jour de la zone cible")
        );
      }

      for (const shape of source) {
        if (shape.id !== targetSnapshot.id) {
          shape.delete();
        }
      }

      try {
        await context.sync();
      } catch (error) {
        throw new Error(
          describeOfficeError(error, "Suppression des zones absorbées")
        );
      }

      if (selectResult) {
        try {
          slide.setSelectedShapes([targetSnapshot.id]);
          await context.sync();
        } catch (error) {
          console.warn(describeOfficeError(error, "Resélection"));
        }
      }

      return targetSnapshot.id;
    });
  }

  function clusterColumns(shapes) {
    const settings = loadSettings();

    if (!settings.respectColumns || shapes.length < 4) {
      return [shapes];
    }

    const fontSize = median(shapes.map(shape => shape.font.size || 10));
    const tolerance = Math.max(10, fontSize * 1.55);
    const ordered = [...shapes].sort(
      (a, b) => a.left - b.left || a.top - b.top
    );
    const clusters = [];

    for (const shape of ordered) {
      let bestIndex = -1;
      let bestDistance = Infinity;

      for (let index = 0; index < clusters.length; index++) {
        const anchor = median(clusters[index].map(item => item.left));
        const distance = Math.abs(shape.left - anchor);

        if (distance <= tolerance && distance < bestDistance) {
          bestIndex = index;
          bestDistance = distance;
        }
      }

      if (bestIndex >= 0) {
        clusters[bestIndex].push(shape);
      } else {
        clusters.push([shape]);
      }
    }

    const meaningful = clusters.filter(cluster => cluster.length >= 2);

    return meaningful.length >= 2 ? meaningful : [shapes];
  }

  function splitParagraphs(lines) {
    const settings = loadSettings();
    const ordered = [...lines].sort(
      (a, b) => a.top - b.top || a.left - b.left
    );

    if (ordered.length <= 1) {
      return ordered.map(line => [line]);
    }

    const pitch = nominalPitch(ordered);
    const groups = [[ordered[0]]];

    for (let index = 1; index < ordered.length; index++) {
      const previous = ordered[index - 1];
      const current = ordered[index];

      const gapBreak =
        current.top - previous.top > pitch * settings.paragraphGap;
      const titleBreak =
        settings.splitOnStyle && styleBreak(previous, current);

      if (gapBreak || titleBreak) {
        groups.push([current]);
      } else {
        groups[groups.length - 1].push(current);
      }
    }

    return groups;
  }

  async function selectResultShapes(slideId, ids) {
    if (!ids.length) return;

    try {
      await PowerPoint.run(async context => {
        const slide = context.presentation.slides.getItem(slideId);
        slide.setSelectedShapes(ids);
        await context.sync();
      });
    } catch (error) {
      console.warn(describeOfficeError(error, "Resélection finale"));
    }
  }

  async function mergeParagraphSelection() {
    const data = await getSelection();

    if (data.shapes.length < 2) {
      throw new Error("Sélectionne au moins deux zones de texte.");
    }

    const columns = clusterColumns(data.shapes);
    const groups = columns.flatMap(column => splitParagraphs(column));
    const mergeable = groups.filter(group => group.length >= 2);
    const untouched = groups
      .filter(group => group.length === 1)
      .flat()
      .map(shape => shape.id);

    if (!mergeable.length) {
      throw new Error(
        "Aucun groupe de lignes à fusionner n’a été détecté."
      );
    }

    const mergedIds = [];

    for (const group of mergeable) {
      mergedIds.push(
        await mergeGroupStable(
          data.slideId,
          group.map(shape => shape.id),
          "paragraph",
          false
        )
      );
    }

    await selectResultShapes(
      data.slideId,
      [...untouched, ...mergedIds]
    );

    setStatus(
      `${mergeable.length} paragraphe(s) fusionné(s) en conservant les zones existantes.`,
      "success"
    );
  }

  async function mergeAllSelection() {
    const data = await getSelection();

    if (data.shapes.length < 2) {
      throw new Error("Sélectionne au moins deux zones de texte.");
    }

    await mergeGroupStable(
      data.slideId,
      data.shapes.map(shape => shape.id),
      "all",
      true
    );

    setStatus(
      `${data.shapes.length} zones fusionnées dans la zone cible.`,
      "success"
    );
  }

  async function mergeColumnsSelection() {
    const data = await getSelection();

    if (data.shapes.length < 4) {
      throw new Error(
        "Sélectionne les lignes d’au moins deux colonnes."
      );
    }

    const columns = clusterColumns(data.shapes);

    if (columns.length < 2) {
      throw new Error(
        "Aucune séparation en colonnes n’a été détectée."
      );
    }

    const resultIds = [];

    for (const column of columns) {
      resultIds.push(
        await mergeGroupStable(
          data.slideId,
          column.map(shape => shape.id),
          "column",
          false
        )
      );
    }

    await selectResultShapes(data.slideId, resultIds);

    setStatus(
      `${columns.length} colonnes fusionnées en ${columns.length} zones distinctes.`,
      "success"
    );
  }

  async function runCommand(event, fn) {
    try {
      await fn();
    } catch (error) {
      await showError(
        error instanceof Error && !error.code
          ? error.message
          : describeOfficeError(error)
      );
    } finally {
      try {
        event.completed();
      } catch {}
    }
  }

  function bindHelpTooltips() {
    const tooltip = document.getElementById("helpTooltip");
    if (!tooltip) return;

    const clamp = (value, min, max) =>
      Math.max(min, Math.min(max, value));

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
      const viewportWidth = document.documentElement.clientWidth;
      const viewportHeight = document.documentElement.clientHeight;

      let left =
        anchor.left + anchor.width / 2 - tip.width / 2;
      left = clamp(
        left,
        8,
        Math.max(8, viewportWidth - tip.width - 8)
      );

      let top = anchor.top - tip.height - margin;

      if (top < 8) {
        top = anchor.bottom + margin;
      }

      if (top + tip.height > viewportHeight - 8) {
        top = clamp(
          anchor.top + anchor.height / 2 - tip.height / 2,
          8,
          Math.max(8, viewportHeight - tip.height - 8)
        );
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
    document
      .querySelector("main")
      ?.addEventListener("scroll", hide, { passive: true });
  }

  function bindSettingsUi() {
    const q = id => document.getElementById(id);
    const settings = loadSettings();

    const gap = q("paragraphGap");
    if (!gap) return;

    const gapValue = q("paragraphGapValue");
    const split = q("splitOnStyle");
    const columns = q("respectColumns");
    const layout = q("layoutMode");
    const blanks = q("maxBlankLines");
    const blanksValue = q("maxBlankLinesValue");
    const dehyphenate = q("dehyphenate");

    gap.value = String(
      Math.round(settings.paragraphGap * 100)
    );
    gapValue.textContent = `${gap.value} %`;
    split.checked = !!settings.splitOnStyle;
    columns.checked = !!settings.respectColumns;
    layout.value = settings.layoutMode;
    blanks.value = String(settings.maxBlankLines);
    blanksValue.textContent = `${blanks.value} lignes`;
    dehyphenate.checked = !!settings.dehyphenate;

    const update = () => {
      saveSettings({
        paragraphGap: Number(gap.value) / 100,
        splitOnStyle: split.checked,
        respectColumns: columns.checked,
        layoutMode: layout.value,
        maxBlankLines: Number(blanks.value),
        dehyphenate: dehyphenate.checked
      });

      gapValue.textContent = `${gap.value} %`;
      blanksValue.textContent = `${blanks.value} lignes`;
      setStatus("Options enregistrées.", "success");
    };

    [
      gap,
      split,
      columns,
      layout,
      blanks,
      dehyphenate
    ].forEach(control =>
      control?.addEventListener("input", update)
    );
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
      Office.actions.associate(
        "mergeColumnsCommand",
        event => runCommand(event, mergeColumnsSelection)
      );
    } catch (error) {
      console.error(error);
    }

    bindSettingsUi();
    bindHelpTooltips();

    const badge = document.getElementById("apiBadge");

    if (badge) {
      const compatible =
        info.host === Office.HostType.PowerPoint &&
        supports("PowerPointApi", "1.5");
      badge.textContent = compatible
        ? "PowerPoint API 1.5+"
        : "API incompatible";
    }

    setStatus("Prêt — mode stable v0.6.");
  });
})();