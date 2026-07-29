#!/usr/bin/env node
// Reproduction for the dead `defaultSize` / `defaultFont` options on
// `toFlowBlocks` (see docs/reports/core-default-typography-options.md).
//
// `ToFlowBlocksOptions` declares `defaultFont` and `defaultSize` as public,
// documented options. The shipped implementation normalizes both onto its
// internal options object and then never reads them, so passing them has no
// effect on the emitted FlowBlocks.
//
// This script proves that behaviorally rather than by reading minified output:
// it converts the same one-paragraph ProseMirror document three ways and
// compares the results.
//
//   A  no options                      -> baseline
//   B  { defaultSize, defaultFont }    -> should differ from A, currently does not
//   C  paragraph attr `defaultTextFormatting` -> the channel that does work
//
// Exits 1 while B is indistinguishable from A (bug present), 0 once the
// options are honored. It is deliberately NOT named *.test.mjs: the repo's
// `bun test` run cannot resolve a real core (the workspace pins the
// contract-only `@docx-editor.dev/core` placeholder), so this has to stay a
// manual script pointed at an installed build.
//
// Usage:
//   node scripts/repro-default-typography-options.mjs
//   node scripts/repro-default-typography-options.mjs --core <path-to-installed-core>
//
// Verified against @eigenpal/docx-editor-core@1.9.0 (dist-tag `latest`).

import { readFileSync, existsSync } from 'node:fs';
import { createRequire } from 'node:module';
import path from 'node:path';
import process from 'node:process';

const CANDIDATE_PACKAGES = ['@docx-editor.dev/core', '@eigenpal/docx-editor-core'];

function coreRootFromArgs() {
  const flag = process.argv.indexOf('--core');
  if (flag === -1) return null;
  const value = process.argv[flag + 1];
  if (!value) fail('--core needs a path to an installed core package directory.');
  const root = path.resolve(value);
  if (!existsSync(path.join(root, 'package.json'))) {
    fail(`No package.json under ${root}.`);
  }
  return root;
}

function coreRootFromResolution() {
  const require = createRequire(path.join(process.cwd(), 'noop.js'));
  for (const name of CANDIDATE_PACKAGES) {
    try {
      return path.dirname(require.resolve(`${name}/package.json`));
    } catch {
      // Try the next candidate.
    }
  }
  return null;
}

function fail(message) {
  console.error(`error: ${message}`);
  process.exit(2);
}

const coreRoot = coreRootFromArgs() ?? coreRootFromResolution();
if (!coreRoot) {
  fail(
    `Could not resolve an installed core package (tried ${CANDIDATE_PACKAGES.join(', ')}).\n` +
      'Point the script at one explicitly:\n' +
      '  node scripts/repro-default-typography-options.mjs --core ' +
      '<node_modules>/@eigenpal/docx-editor-core'
  );
}

const pkg = JSON.parse(readFileSync(path.join(coreRoot, 'package.json'), 'utf8'));

/** Resolve an export subpath to an absolute ESM entry using the package's own export map. */
function entry(subpath) {
  const declaration = pkg.exports?.[subpath];
  if (!declaration) fail(`${pkg.name}@${pkg.version} does not export "${subpath}".`);
  const target =
    typeof declaration === 'string'
      ? declaration
      : (declaration.import?.default ?? declaration.import ?? declaration.default);
  if (typeof target !== 'string') fail(`Could not resolve an ESM entry for "${subpath}".`);
  return path.join(coreRoot, target);
}

const { toFlowBlocks } = await import(entry('./layout-bridge/toFlowBlocks'));
const { schema } = await import(entry('./prosemirror'));

/** One paragraph, one text node, no character marks — so nothing overrides the default. */
function oneParagraph(paragraphAttrs = {}) {
  return schema.nodes.doc.create(null, [
    schema.nodes.paragraph.create(paragraphAttrs, [schema.text('Hello')]),
  ]);
}

/** The typography actually carried by the emitted blocks, stripped of positions and ids. */
function typography(blocks) {
  return blocks.map((block) => ({
    kind: block.kind,
    defaultFontFamily: block.defaultFontFamily,
    defaultFontSize: block.defaultFontSize,
    runs: (block.runs ?? []).map((run) => ({
      text: run.text,
      fontFamily: run.fontFamily,
      fontSize: run.fontSize,
    })),
  }));
}

const cases = {
  A: typography(toFlowBlocks(oneParagraph())),
  B: typography(toFlowBlocks(oneParagraph(), { defaultSize: 24, defaultFont: 'Arial' })),
  C: typography(
    toFlowBlocks(
      oneParagraph({
        // What the parser produces from `docDefaults.rPr` in word/styles.xml.
        // `fontSize` is in half-points, per ECMA-376 §17.3.2.38 (`w:sz`).
        defaultTextFormatting: { fontSize: 20, fontFamily: { ascii: 'Arial', hAnsi: 'Arial' } },
      })
    )
  ),
};

console.log(`core package: ${pkg.name}@${pkg.version}`);
console.log(`resolved from: ${coreRoot}\n`);
console.log('A  no options');
console.log(JSON.stringify(cases.A, null, 2));
console.log('\nB  toFlowBlocks(doc, { defaultSize: 24, defaultFont: "Arial" })');
console.log(JSON.stringify(cases.B, null, 2));
console.log('\nC  paragraph attr defaultTextFormatting { fontSize: 20 half-points, fontFamily }');
console.log(JSON.stringify(cases.C, null, 2));

const optionsIgnored = JSON.stringify(cases.A) === JSON.stringify(cases.B);
const attrChannelWorks = JSON.stringify(cases.A) !== JSON.stringify(cases.C);

console.log('');
if (optionsIgnored) {
  console.log('FAIL  B is byte-identical to A: `defaultSize` and `defaultFont` are ignored.');
  if (attrChannelWorks) {
    console.log(
      '      C does change the output, so the paragraph-attr channel fed by ' +
        '`docDefaults.rPr` is the only working way to set document default typography.'
    );
  }
  process.exit(1);
}

console.log('PASS  `defaultSize` / `defaultFont` reached the emitted blocks.');
