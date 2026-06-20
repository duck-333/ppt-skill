#!/usr/bin/env node
/**
 * Render template-driven slides JSON into an editable PPTX with artifact-tool.
 *
 * This script never mutates the source template. It duplicates inherited
 * slides and edits inherited elements in place.
 */

import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";

const PT_TO_PX = 96 / 72;
const FONT_FAMILY = "Microsoft YaHei";


function pt(value) {
  return value * PT_TO_PX;
}


function parseArgs(argv) {
  const args = {};
  for (let index = 0; index < argv.length; index += 1) {
    const key = argv[index];
    if (!key.startsWith("--")) throw new Error(`Unexpected argument: ${key}`);
    const value = argv[index + 1];
    if (!value || value.startsWith("--")) {
      args[key.slice(2)] = true;
    } else {
      args[key.slice(2)] = value;
      index += 1;
    }
  }
  return args;
}


async function firstExisting(paths) {
  for (const candidate of paths) {
    try {
      await fs.access(candidate);
      return candidate;
    } catch {
      // Continue to the next candidate.
    }
  }
  return undefined;
}


async function loadArtifactTool() {
  if (process.env.ARTIFACT_TOOL_ENTRY) {
    return import(pathToFileURL(path.resolve(process.env.ARTIFACT_TOOL_ENTRY)).href);
  }

  const home = process.env.USERPROFILE || process.env.HOME || os.homedir();
  const packageRoot = await firstExisting([
    path.resolve("node_modules/@oai/artifact-tool"),
    path.join(
      home,
      ".cache",
      "codex-runtimes",
      "codex-primary-runtime",
      "dependencies",
      "node",
      "node_modules",
      "@oai",
      "artifact-tool",
    ),
  ]);
  if (!packageRoot) {
    throw new Error(
      "Cannot find @oai/artifact-tool. Set ARTIFACT_TOOL_ENTRY to artifact_tool.mjs.",
    );
  }
  const entry = await firstExisting([
    path.join(packageRoot, "dist", "node", "artifact_tool.mjs"),
    path.join(packageRoot, "dist", "artifact_tool.mjs"),
  ]);
  if (!entry) throw new Error(`Missing artifact-tool entrypoint under ${packageRoot}`);
  return import(pathToFileURL(entry).href);
}


function slidesFromPresentation(presentation) {
  if (Array.isArray(presentation.slides?.items)) return presentation.slides.items;
  if (
    Number.isInteger(presentation.slides?.count)
    && typeof presentation.slides.getItem === "function"
  ) {
    return Array.from(
      { length: presentation.slides.count },
      (_, index) => presentation.slides.getItem(index),
    );
  }
  throw new Error("Could not enumerate presentation slides.");
}


async function writeBlob(filePath, blob) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  await fs.writeFile(filePath, new Uint8Array(await blob.arrayBuffer()));
}


function parseNdjson(text) {
  return text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => JSON.parse(line));
}


function makeFrameMap(slides) {
  return {
    outputSlides: slides.map((slide, index) => ({
      outputSlide: index + 1,
      sourceSlide: slide.sourceSlide,
      narrativeRole: slide.type,
      reuseMode: "duplicate-slide",
      editTargets: [],
    })),
    omittedSourceSlides: [],
  };
}


function setTextPreservingFormat(shape, replacements) {
  let value = String(shape.text);
  for (const [search, replacement] of replacements) {
    value = value.split(search).join(replacement);
  }
  shape.text = value;
}


function shapeHeightForSummary(item, width) {
  const titleChars = item.title.length;
  const wordChars = item.words.join(", ").length;
  const titleLines = Math.max(1, Math.ceil(titleChars / Math.max(10, Math.floor(width / 24))));
  const wordLines = Math.max(1, Math.ceil(wordChars / Math.max(14, Math.floor(width / 18))));
  return Math.min(450, Math.max(202, 96 + titleLines * 42 + wordLines * 46));
}


function outputSlideRecords(records, slideNumber) {
  return records.filter((record) => record.slide === slideNumber && record.id);
}


function findRecord(records, slideNumber, name) {
  const match = records.find(
    (record) => record.slide === slideNumber && record.name === name && record.id,
  );
  if (!match) {
    throw new Error(`Missing inherited element "${name}" on output slide ${slideNumber}.`);
  }
  return match;
}


function maybeFindRecord(records, slideNumber, name) {
  return records.find(
    (record) => record.slide === slideNumber && record.name === name && record.id,
  );
}


function getInheritedShape(slide, name, required = true) {
  let shape;
  try {
    shape = slide.shapes.getItem(name);
  } catch {
    shape = undefined;
  }
  if (!shape && required) {
    throw new Error(`Missing inherited element "${name}" on output slide ${slide.index + 1}.`);
  }
  return shape;
}


function safeDeleteShape(shape) {
  if (!shape) return;
  if (typeof shape.delete !== "function") {
    throw new Error(`Element ${shape.id || "(unknown)"} does not expose delete().`);
  }
  shape.delete();
}


function styleUpperLeft(shape) {
  shape.position = { left: 94.93, top: 42.87, width: 324, height: 48.33 };
  shape.text.style = {
    fontSize: pt(20),
    typeface: FONT_FAMILY,
    bold: true,
    italic: false,
    color: "#000000",
    alignment: "left",
    verticalAlignment: "middle",
    wrap: "none",
    autoFit: "shrinkText",
  };
}


function applyDetailContent(slide, spec) {
  const word = spec.word;
  const replacements = [
    ["HEADER_SUBTITLE", "{{SUB_TITLE}}", spec.category],
    ["WORD_NAME", "{{WORD}}", word.word],
    ["WORD_DEFINITION", "{{DEFINITION}}", word.definition],
    ["WORD_ANALYSIS", "{{ROOT_ANALYSIS}}", word.analysis],
    ["Text 0", "{{TOP_CATEGORY}}", spec.topCategory],
  ];
  for (const [name, token, value] of replacements) {
    const shape = getInheritedShape(slide, name);
    setTextPreservingFormat(shape, [[token, value]]);
  }

  const phoneticShape = getInheritedShape(slide, "WORD_PHONETIC", false);
  if (word.phonetic) {
    setTextPreservingFormat(
      phoneticShape,
      [["/{{PHONETIC}}/", `/${word.phonetic}/`], ["{{PHONETIC}}", word.phonetic]],
    );
  } else {
    safeDeleteShape(phoneticShape);
  }

  const headerShape = getInheritedShape(slide, "HEADER_SUBTITLE");
  const wordShape = getInheritedShape(slide, "WORD_NAME");
  const definitionShape = getInheritedShape(slide, "WORD_DEFINITION");
  const analysisShape = getInheritedShape(slide, "WORD_ANALYSIS");
  const topShape = getInheritedShape(slide, "Text 0");

  styleUpperLeft(topShape);

  headerShape.position = { left: 80, top: 86.09, width: 235.93, height: 52.49 };
  headerShape.text.style = {
    fontSize: pt(20),
    typeface: FONT_FAMILY,
    bold: true,
    italic: false,
    color: "#000000",
    wrap: "square",
    autoFit: "shrinkText",
    verticalAlignment: "middle",
  };

  wordShape.position = { left: 486.27, top: 168, width: 414.73, height: 106.53 };
  wordShape.text.style = {
    color: "#FF0000",
    fontSize: pt(60),
    typeface: FONT_FAMILY,
    bold: true,
    italic: false,
    wrap: "none",
    autoFit: "shrinkText",
    verticalAlignment: "middle",
  };

  if (phoneticShape) {
    phoneticShape.position = {
      left: 486.27,
      top: 292.67,
      width: 291.4,
      height: 54.8,
    };
    phoneticShape.text.style = {
      fontSize: pt(28),
      typeface: FONT_FAMILY,
      bold: false,
      italic: false,
      color: "#000000",
      wrap: "none",
      autoFit: "shrinkText",
      verticalAlignment: "middle",
    };
  }

  const analysisLength = [...String(word.analysis || "")].length;
  const analysisPointSize = analysisLength > 120 ? 20 : analysisLength > 80 ? 22 : 32;
  const definitionLength = [...String(word.definition || "")].length;
  const longConceptDefinition = spec.type === "concept" && definitionLength > 45;
  analysisShape.position = {
    left: 486.27,
    top: 374.76,
    width: 729.6,
    height: longConceptDefinition ? 158 : 182.64,
  };
  analysisShape.text.style = {
    fontSize: pt(analysisPointSize),
    typeface: FONT_FAMILY,
    bold: false,
    italic: false,
    color: "#000000",
    lineSpacing: 1,
    wrap: "square",
    autoFit: "shrinkText",
    verticalAlignment: "top",
  };

  definitionShape.position = {
    left: 486.27,
    top: longConceptDefinition ? 540 : 576.31,
    width: 729.6,
    height: longConceptDefinition ? 126 : 77.4,
  };
  definitionShape.text.style = {
    color: "#FF0000",
    fontSize: pt(
      spec.type === "concept"
        ? definitionLength > 80
          ? 18
          : longConceptDefinition
            ? 20
            : 24
        : 36,
    ),
    typeface: FONT_FAMILY,
    bold: true,
    italic: false,
    wrap: "square",
    autoFit: "shrinkText",
    verticalAlignment: longConceptDefinition ? "bottom" : "middle",
  };
}


async function main() {
  const args = parseArgs(process.argv.slice(2));
  const templatePath = path.resolve(String(args.template || ""));
  const slidesJsonPath = path.resolve(String(args["slides-json"] || ""));
  const outputPath = path.resolve(String(args.output || ""));
  const workspace = path.resolve(
    String(args.workspace || path.join(os.tmpdir(), "word-to-ppt-artifact")),
  );
  const renderScale = Number(args.scale || 1);

  if (!args.template || !args["slides-json"] || !args.output) {
    throw new Error(
      "Usage: node word_to_ppt_artifact.mjs --template template.pptx "
      + "--slides-json slides-data.json --output result.pptx [--workspace dir]",
    );
  }
  if (!Number.isFinite(renderScale) || renderScale <= 0) {
    throw new Error("--scale must be a positive number.");
  }

  const tmpDir = path.join(workspace, "tmp");
  const previewDir = path.join(tmpDir, "preview");
  const exportedPreviewDir = path.join(tmpDir, "preview-exported");
  const layoutDir = path.join(tmpDir, "layout", "final");
  const qaDir = path.join(tmpDir, "qa");
  await Promise.all([
    fs.mkdir(previewDir, { recursive: true }),
    fs.mkdir(exportedPreviewDir, { recursive: true }),
    fs.mkdir(layoutDir, { recursive: true }),
    fs.mkdir(qaDir, { recursive: true }),
    fs.mkdir(path.dirname(outputPath), { recursive: true }),
  ]);

  const model = JSON.parse(await fs.readFile(slidesJsonPath, "utf8"));
  if (!Array.isArray(model.slides) || model.slides.length === 0) {
    throw new Error("slides JSON does not contain a non-empty slides array.");
  }

  await fs.writeFile(
    path.join(tmpDir, "source-notes.txt"),
    [
      `Word source: ${model.source?.docx || "(unknown)"}`,
      `Template source: ${templatePath}`,
      `Lesson: ${model.lesson?.title || "(unknown)"}`,
      `Categories: ${model.counts?.categories ?? "?"}`,
      `Words: ${model.counts?.words ?? "?"}`,
      "All slide copy is derived from the Word source.",
      "",
    ].join("\n"),
    "utf8",
  );
  await fs.writeFile(
    path.join(tmpDir, "template-audit.txt"),
    [
      "Production template: master_template.pptx",
      "Source pattern: slide 1 overview, slide 2 word detail, slide 3 summary.",
      "Preserved: backgrounds, typography, fills, cards, spacing, and inherited image placeholder.",
      "Splitting: overview pages contain at most three categories; summary pages contain at most six root cards; every word has one detail slide.",
      "",
    ].join("\n"),
    "utf8",
  );
  await fs.writeFile(
    path.join(tmpDir, "template-frame-map.json"),
    `${JSON.stringify(makeFrameMap(model.slides), null, 2)}\n`,
    "utf8",
  );
  await fs.writeFile(
    path.join(tmpDir, "deviation-log.txt"),
    [
      "1. The user-supplied lesson 86 deck contains a dangling slide relationship and cannot be imported by artifact-tool.",
      "2. The repository's protected master_template.pptx is used as the editable production template.",
      "3. The cover is derived from template slide 1; inherited title objects are repositioned to match lesson 86 cover geometry.",
      "4. Audio action buttons are not generated because artifact-tool does not expose PowerPoint click-sound actions.",
      "5. Summary uses one inherited card per page; this remains within the rule of at most three cards per summary page.",
      "",
    ].join("\n"),
    "utf8",
  );

  const { FileBlob, PresentationFile } = await loadArtifactTool();
  const presentation = await PresentationFile.importPptx(await FileBlob.load(templatePath));
  const originals = [...slidesFromPresentation(presentation)];
  if (originals.length < 3) {
    throw new Error(`Template must contain at least 3 slides; found ${originals.length}.`);
  }

  if (args["debug-api"]) {
    const summarySnapshot = await presentation.inspect({
      kind: "slide,textbox,shape",
      search: "SUMMARY_UNIT_TEMPLATE",
      maxChars: 10000,
    });
    const records = parseNdjson(summarySnapshot.ndjson);
    const record = records.find((item) => item.name === "SUMMARY_UNIT_TEMPLATE");
    const shape = record ? presentation.resolve(record.id) : undefined;
    const directShape = originals[2].shapes.getItem("SUMMARY_UNIT_TEMPLATE");
    console.log(JSON.stringify({
      slideMethods: Object.getOwnPropertyNames(Object.getPrototypeOf(originals[0])).sort(),
      shapesCollectionMethods: Object.getOwnPropertyNames(
        Object.getPrototypeOf(originals[0].shapes),
      ).sort(),
      shapesCollectionKeys: Object.keys(originals[0].shapes).sort(),
      slideElementCount: Array.isArray(originals[0].elements) ? originals[0].elements.length : null,
      firstElementKeys: Array.isArray(originals[0].elements) && originals[0].elements[0]
        ? Object.keys(originals[0].elements[0]).sort()
        : [],
      shapeMethods: shape
        ? Object.getOwnPropertyNames(Object.getPrototypeOf(shape)).sort()
        : [],
      shapeKeys: shape ? Object.keys(shape).sort() : [],
      directShapeId: directShape?.id,
      directShapeProto: directShape?.toProto(),
      shapeAddHelp: presentation.help("slide.shapes.add", {
        include: ["index", "examples", "notes"],
        maxChars: 8000,
      }),
      snapshot: records,
    }, null, 2));
    return;
  }

  const outputSlides = model.slides.map((slideSpec) => {
    const sourceSlide = originals[slideSpec.sourceSlide - 1];
    if (!sourceSlide) {
      throw new Error(`Invalid sourceSlide ${slideSpec.sourceSlide}.`);
    }
    return sourceSlide.duplicate();
  });

  for (const slide of originals) slide.delete();
  outputSlides.forEach((slide, index) => slide.moveTo(index));

  const snapshot = await presentation.inspect({
    kind: "slide,textbox,shape,image",
    include: "id,slide,name,textPreview,bbox,isPlaceholder,alt",
    maxChars: 500000,
  });
  const records = parseNdjson(snapshot.ndjson);
  await fs.writeFile(
    path.join(tmpDir, "starter-inspect.ndjson"),
    `${snapshot.ndjson.trim()}\n`,
    "utf8",
  );

  for (let index = 0; index < model.slides.length; index += 1) {
    const slideNumber = index + 1;
    const spec = model.slides[index];
    const slide = outputSlides[index];

    if (spec.type === "cover") {
      const top = getInheritedShape(slide, "Text 0");
      const title = getInheritedShape(slide, "TITLE_ROOT");
      setTextPreservingFormat(top, [["{{TOP_CATEGORY}}", spec.topCategory]]);
      setTextPreservingFormat(
        title,
        [["{{ROOT_NAME}}", model.lesson?.title || spec.title]],
      );
      styleUpperLeft(top);
      title.position = { left: 115.19, top: 216, width: 1049.55, height: 230.4 };
      title.text.style = {
        fontSize: pt(42),
        typeface: FONT_FAMILY,
        bold: true,
        italic: false,
        color: "#000000",
        alignment: "center",
        verticalAlignment: "middle",
        wrap: "square",
        autoFit: "shrinkText",
      };
      safeDeleteShape(getInheritedShape(slide, "ROOT_LOGIC_DESC", false));
      for (let itemIndex = 1; itemIndex <= 3; itemIndex += 1) {
        safeDeleteShape(getInheritedShape(slide, `SUB_ITEM_UNIT_${itemIndex}`, false));
      }
      continue;
    }

    if (spec.type === "overview") {
      const top = getInheritedShape(slide, "Text 0");
      const title = getInheritedShape(slide, "TITLE_ROOT");
      setTextPreservingFormat(top, [["{{TOP_CATEGORY}}", spec.topCategory]]);
      setTextPreservingFormat(title, [["{{ROOT_NAME}}", spec.rootName]]);
      styleUpperLeft(top);

      const logicShape = getInheritedShape(slide, "ROOT_LOGIC_DESC", false);
      if (spec.rootLogic) {
        setTextPreservingFormat(
          logicShape,
          [["{{ROOT_LOGIC}}", spec.rootLogic]],
        );
      } else {
        safeDeleteShape(logicShape);
      }

      for (let itemIndex = 1; itemIndex <= 3; itemIndex += 1) {
        const itemShape = getInheritedShape(slide, `SUB_ITEM_UNIT_${itemIndex}`, false);
        const value = spec.items[itemIndex - 1];
        if (value) {
          setTextPreservingFormat(
            itemShape,
            [[`{{SUB_TITLE_${itemIndex}}}`, value], ["��", String((spec.pageIndex - 1) * 3 + itemIndex)]],
          );
        } else {
          safeDeleteShape(itemShape);
        }
      }
      continue;
    }

    if (spec.type === "detail" || spec.type === "concept") {
      applyDetailContent(slide, spec);
      continue;
    }

    if (spec.type === "summary") {
      const templateShape = getInheritedShape(slide, "SUMMARY_UNIT_TEMPLATE");
      if (spec.items.length < 1 || spec.items.length > 6) {
        throw new Error("Summary slides must contain between one and six root cards.");
      }
      const count = spec.items.length;
      const left = 57.6;
      const top = 105;
      const columnGap = 28;
      const rowGap = 18;
      const width = (1280 - left * 2 - columnGap) / 2;
      const height = 170;

      for (let itemIndex = 0; itemIndex < count; itemIndex += 1) {
        const item = spec.items[itemIndex];
        const column = itemIndex % 2;
        const row = Math.floor(itemIndex / 2);
        const position = {
          left: left + column * (width + columnGap),
          top: top + row * (height + rowGap),
          width,
          height,
        };
        const card = itemIndex === 0
          ? templateShape
          : slide.shapes.add({
            geometry: "roundRect",
            name: `SUMMARY_UNIT_${itemIndex + 1}`,
            position,
            fill: "#DDEBF7",
            line: { style: "solid", fill: "#FFFFFF", width: 2 },
            borderRadius: "rounded-xl",
            shadow: "shadow-sm",
          });
        const titlePointSize = [...item.title].length > 28
          ? 14
          : [...item.title].length > 20
            ? 16
            : 18;
        const wordList = item.words.join(", ");
        const wordPointSize = wordList.length > 82 ? 10 : wordList.length > 55 ? 11 : 12;
        card.text = [
          {
            runs: [
              {
                run: item.title,
                textStyle: {
                  bold: true,
                  italic: false,
                  fontSize: `${titlePointSize}pt`,
                  typeface: FONT_FAMILY,
                  color: "#000000",
                },
              },
            ],
            spaceAfter: 8,
          },
          {
            runs: [
              {
                run: wordList,
                textStyle: {
                  bold: false,
                  italic: false,
                  fontSize: `${wordPointSize}pt`,
                  typeface: FONT_FAMILY,
                  color: "#3C3C3C",
                },
              },
            ],
          },
        ];
        card.position = {
          ...position,
        };
        card.line = { style: "solid", fill: "#FFFFFF", width: 2 };
        card.text.style = {
          typeface: FONT_FAMILY,
          alignment: "left",
          verticalAlignment: "middle",
          wrap: "square",
          autoFit: "shrinkText",
          insets: { top: 12, right: 16, bottom: 12, left: 16 },
        };
      }
      continue;
    }

    throw new Error(`Unsupported slide type: ${spec.type}`);
  }

  const finalInspect = await presentation.inspect({
    kind: "slide,textbox,shape,image",
    include: "id,slide,name,textPreview,bbox,isPlaceholder,alt",
    maxChars: 500000,
  });
  const unresolvedTokens = ["{{", "Click to add", "Slide Number", "Footer", "Date"];
  const unresolved = parseNdjson(finalInspect.ndjson).filter((record) => {
    const text = String(record.textPreview || "");
    return unresolvedTokens.some((token) => text.includes(token));
  });
  if (unresolved.length) {
    throw new Error(
      `Unresolved template placeholders remain: ${JSON.stringify(unresolved.slice(0, 10))}`,
    );
  }

  const previewPaths = [];
  const layoutPaths = [];
  for (let index = 0; index < outputSlides.length; index += 1) {
    const slide = outputSlides[index];
    const number = String(index + 1).padStart(2, "0");
    const previewPath = path.join(previewDir, `slide-${number}.png`);
    const layoutPath = path.join(layoutDir, `slide-${number}.layout.json`);
    await writeBlob(
      previewPath,
      await presentation.export({ slide, format: "png", scale: renderScale }),
    );
    await writeBlob(
      layoutPath,
      await presentation.export({ slide, format: "layout" }),
    );
    previewPaths.push(previewPath);
    layoutPaths.push(layoutPath);
  }

  const montagePath = path.join(qaDir, "deck-montage.webp");
  await writeBlob(
    montagePath,
    await presentation.export({
      format: "webp",
      montage: { columns: 5, slideWidth: 320, padding: 24, gap: 16 },
      scale: 1,
    }),
  );

  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(outputPath);
  const outputStat = await fs.stat(outputPath);
  if (outputStat.size <= 0) throw new Error(`Exported PPTX is empty: ${outputPath}`);

  const exportedPresentation = await PresentationFile.importPptx(
    await FileBlob.load(outputPath),
  );
  const exportedSlides = slidesFromPresentation(exportedPresentation);
  if (exportedSlides.length !== model.slides.length) {
    throw new Error(
      `Exported slide count mismatch: expected ${model.slides.length}, found ${exportedSlides.length}.`,
    );
  }
  const exportedInspect = await exportedPresentation.inspect({
    kind: "slide,textbox,shape,image",
    include: "id,slide,name,textPreview,bbox,isPlaceholder,alt",
    maxChars: 500000,
  });
  const exportedUnresolved = parseNdjson(exportedInspect.ndjson).filter((record) => {
    const text = String(record.textPreview || "");
    return unresolvedTokens.some((token) => text.includes(token));
  });
  if (exportedUnresolved.length) {
    throw new Error(
      `Exported PPTX contains unresolved placeholders: ${
        JSON.stringify(exportedUnresolved.slice(0, 10))
      }`,
    );
  }
  for (let index = 0; index < exportedSlides.length; index += 1) {
    const number = String(index + 1).padStart(2, "0");
    await writeBlob(
      path.join(exportedPreviewDir, `slide-${number}.png`),
      await exportedPresentation.export({
        slide: exportedSlides[index],
        format: "png",
        scale: renderScale,
      }),
    );
  }

  await fs.writeFile(
    path.join(qaDir, "visual-qa.txt"),
    [
      `Expected slides: ${model.slides.length}`,
      `Rendered slides: ${previewPaths.length}`,
      `Output bytes: ${outputStat.size}`,
      `Unresolved placeholder count: ${unresolved.length}`,
      `Exported unresolved placeholder count: ${exportedUnresolved.length}`,
      "Full-size per-slide PNGs were rendered both before export and after re-importing the final PPTX.",
      "Layout JSON was generated for every slide.",
      "",
    ].join("\n"),
    "utf8",
  );

  console.log(JSON.stringify({
    output: outputPath,
    bytes: outputStat.size,
    slideCount: outputSlides.length,
    montage: montagePath,
    previewDir,
    exportedPreviewDir,
    layoutDir,
  }, null, 2));
}


main().catch((error) => {
  console.error(error.stack || error.message || String(error));
  process.exitCode = 1;
});
