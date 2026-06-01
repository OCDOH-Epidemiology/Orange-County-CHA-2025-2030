#!/usr/bin/env node
"use strict";

const fs = require("fs");
const path = require("path");

const repoRoot = path.resolve(__dirname, "..");
const candidateTargetPaths = [
  path.join(repoRoot, "site_libs", "quarto-search", "quarto-search.js"),
  path.join(
    repoRoot,
    "docs",
    "site_libs",
    "quarto-search",
    "quarto-search.js"
  ),
];

function replaceOnce(contents, before, after, label) {
  if (contents.includes(after)) {
    return contents;
  }
  if (!contents.includes(before)) {
    throw new Error(`Patch anchor not found for ${label}`);
  }
  return contents.replace(before, after);
}

function replaceAny(contents, befores, after, label) {
  if (contents.includes(after)) {
    return contents;
  }
  for (const before of befores) {
    if (contents.includes(before)) {
      return contents.replace(before, after);
    }
  }
  throw new Error(`Patch anchor not found for ${label}`);
}

function patchSearchRuntime(contents) {
  let updated = contents;

  const highlightBeforeOriginal = `  // highlight matches on the page
  if (query && mainEl) {
    // perform any highlighting
    highlight(escapeRegExp(query), mainEl);

    // fix up the URL to remove the q query param
    const replacementUrl = new URL(window.location);
    replacementUrl.searchParams.delete(kQueryArg);
    window.history.replaceState({}, "", replacementUrl);
  }`;

  const highlightBeforePatchedV1 = `  // highlight matches on the page
  if (query && mainEl) {
    // perform any highlighting
    highlight(escapeRegExp(query), mainEl);

    // CHA PATCH: jump directly to first highlighted hit in the loaded page.
    // This ensures search clicks land on the actual matching text.
    const firstMatch = mainEl.querySelector("mark");
    if (firstMatch) {
      firstMatch.scrollIntoView({ block: "center", inline: "nearest" });
    }

    // fix up the URL to remove the q query param
    const replacementUrl = new URL(window.location);
    replacementUrl.searchParams.delete(kQueryArg);
    window.history.replaceState({}, "", replacementUrl);
  }`;

  const highlightAfter = `  // highlight matches on the page
  if (query && mainEl) {
    // perform exact phrase highlighting first
    const marksBefore = mainEl.querySelectorAll("mark").length;
    highlight(escapeRegExp(query), mainEl);
    const marksAfter = mainEl.querySelectorAll("mark").length;

    // CHA PATCH: if exact phrase doesn't produce a visible hit, fallback
    // to highlighting significant individual terms from the query.
    if (marksAfter === marksBefore) {
      const terms = [...new Set(query.split(/\\s+/))]
        .map((term) => term.trim())
        .filter((term) => term.length >= 3);
      terms.forEach((term) => {
        highlight(escapeRegExp(term), mainEl);
      });
    }

    // CHA PATCH: jump directly to first highlighted hit in the loaded page.
    const firstMatch = mainEl.querySelector("mark");
    if (firstMatch) {
      firstMatch.scrollIntoView({ block: "center", inline: "nearest" });
    }

    // fix up the URL to remove the q query param
    const replacementUrl = new URL(window.location);
    replacementUrl.searchParams.delete(kQueryArg);
    window.history.replaceState({}, "", replacementUrl);
  }`;

  updated = replaceAny(
    updated,
    [highlightBeforeOriginal, highlightBeforePatchedV1],
    highlightAfter,
    "scroll to first highlighted match"
  );

  const resetBefore = `  // Clear search highlighting when the user scrolls sufficiently
  const resetFn = () => {
    resetHighlighting("");
    window.removeEventListener("quarto-hrChanged", resetFn);
    window.removeEventListener("quarto-sectionChanged", resetFn);
  };

  // Register this event after the initial scrolling and settling of events
  // on the page
  window.addEventListener("quarto-hrChanged", resetFn);
  window.addEventListener("quarto-sectionChanged", resetFn);`;

  const resetAfter = `  // CHA PATCH: keep highlight visible after landing from search.
  // Do not clear on initial section change; clear on first user interaction.
  const resetFn = () => {
    window.removeEventListener("quarto-hrChanged", resetFn);
    window.removeEventListener("quarto-sectionChanged", resetFn);
  };

  // Register this event after the initial scrolling and settling of events
  // on the page
  window.addEventListener("quarto-hrChanged", resetFn);
  window.addEventListener("quarto-sectionChanged", resetFn);

  const clearOnInteraction = () => {
    resetHighlighting("");
    window.removeEventListener("pointerdown", clearOnInteraction, true);
    window.removeEventListener("keydown", clearOnInteraction, true);
  };
  window.addEventListener("pointerdown", clearOnInteraction, true);
  window.addEventListener("keydown", clearOnInteraction, true);`;

  updated = replaceOnce(
    updated,
    resetBefore,
    resetAfter,
    "persist highlight until interaction"
  );

  const reshapeBefore = `          const firstItem = value[0];
            reshapedItems.push({
              ...firstItem,
              type: kItemTypeDoc,
            });`;

  const reshapeAfter = `          // CHA PATCH: prefer section anchors for top-level document links.
          const firstItem = value[0];
            const anchorItem = value.find((item) => item.href.includes("#"));
            const preferredHref = anchorItem ? anchorItem.href : firstItem.href;
            reshapedItems.push({
              ...firstItem,
              href: preferredHref,
              type: kItemTypeDoc,
            });`;

  updated = replaceOnce(
    updated,
    reshapeBefore,
    reshapeAfter,
    "anchor-first top result links"
  );

  const helperBefore = `let subSearchTerm = undefined;
let subSearchFuse = undefined;
const kFuseMaxWait = 125;

async function fuseSearch(query, fuse, fuseOptions) {`;

  const helperAfter = `let subSearchTerm = undefined;
let subSearchFuse = undefined;
const kFuseMaxWait = 125;

function chapterOrderFromCrumbs(crumbs) {
  if (!crumbs || crumbs.length === 0) {
    return Number.POSITIVE_INFINITY;
  }
  const chapterMatch = crumbs[0].match(/chapter-number[^>]*>(\\d+)/);
  if (!chapterMatch) {
    return Number.POSITIVE_INFINITY;
  }
  const order = Number(chapterMatch[1]);
  return Number.isFinite(order) ? order : Number.POSITIVE_INFINITY;
}

async function fuseSearch(query, fuse, fuseOptions) {`;

  updated = replaceOnce(
    updated,
    helperBefore,
    helperAfter,
    "chapter order helper"
  );

  const searchBefore = `  // Search using the active fuse
  const then = performance.now();
  const resultsRaw = await index.search(query, fuseOptions);
  const now = performance.now();

  const results = resultsRaw.map((result) => {`;

  const searchAfter = `  // Search using the active fuse
  const then = performance.now();
  const resultsRaw = await index.search(query, fuseOptions);
  const now = performance.now();

  // CHA PATCH: force chronological chapter ordering while preserving
  // Fuse relevance order inside each chapter.
  const sortedResultsRaw = resultsRaw
    .map((result, originalIndex) => ({ result, originalIndex }))
    .sort((a, b) => {
      const chapterDelta =
        chapterOrderFromCrumbs(a.result.item.crumbs) -
        chapterOrderFromCrumbs(b.result.item.crumbs);
      if (chapterDelta !== 0) {
        return chapterDelta;
      }
      return a.originalIndex - b.originalIndex;
    })
    .map((entry) => entry.result);

  const results = sortedResultsRaw.map((result) => {`;

  updated = replaceOnce(
    updated,
    searchBefore,
    searchAfter,
    "chapter-first fuse sort"
  );

  return updated;
}

function main() {
  const existingTargets = candidateTargetPaths.filter((targetPath) =>
    fs.existsSync(targetPath)
  );
  if (existingTargets.length === 0) {
    console.log(
      "Skipping quarto search patch; generated search runtime not found in expected locations: " +
        candidateTargetPaths.join(", ")
    );
    return;
  }

  let changedCount = 0;
  for (const targetPath of existingTargets) {
    const original = fs.readFileSync(targetPath, "utf8");
    const patched = patchSearchRuntime(original);
    if (patched !== original) {
      fs.writeFileSync(targetPath, patched, "utf8");
      changedCount += 1;
      console.log(`Patched ${path.relative(repoRoot, targetPath)}`);
    } else {
      console.log(
        `No patch changes needed (already applied): ${path.relative(
          repoRoot,
          targetPath
        )}`
      );
    }
  }
  if (changedCount === 0) {
    console.log("Patch run complete: no files changed.");
  }
}

main();
