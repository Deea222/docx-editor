# `toFlowBlocks` ignores its documented `defaultSize` and `defaultFont` options

## Summary

`ToFlowBlocksOptions` declares `defaultFont` and `defaultSize` as public options, each with a doc comment.
The shipped implementation normalizes both onto its internal options object and then never reads either one, so passing them has no effect on the emitted `FlowBlock[]`.
There is consequently no supported way to set the document default font size or family: when a run carries no explicit size, the layout engine falls back to a hardcoded 11pt Calibri.

This is an API bug, not a security issue.
The practical cost is that a documented knob silently does nothing, and callers who need a different default have to reach for an undocumented workaround.

Affected version: `@eigenpal/docx-editor-core@1.9.0`, the current `latest` dist-tag (the package is flagged deprecated on npm).
The same options are still declared in the type at the time of writing, and no CHANGELOG entry through `1.10.0` addresses them.

## What the type promises

`dist/layout-bridge/toFlowBlocks.d.ts`:

```ts
/**
 * Options for the conversion.
 */
type ToFlowBlocksOptions = {
  /** Default font family. */
  defaultFont?: string;
  /** Default font size in points. */
  defaultSize?: number;
  /** Theme for resolving theme colors. */
  theme?: Theme | null;
  ...
};
```

The file's `@packageDocumentation` block marks the deep import as intentionally public:

> The deep import `@eigenpal/.../layout-bridge/toFlowBlocks` is part of the public surface (Vue adapter + tests) [...]

So this is not an incidental export.
Both options are `@public`, named, and documented.

## What the implementation does

In the minified ESM build, `toFlowBlocks` is `function xe(e, s = {})` in `dist/chunk-5UCYRCUV.mjs`.
Its first statement normalizes the options:

```js
function xe(e, s = {}) {
  let i = e.attrs?.defaultTabStopTwips,
    n = {
      ...s,
      defaultFont: s.defaultFont ?? J,
      defaultSize: s.defaultSize ?? Q,
      defaultTabStopTwips: s.defaultTabStopTwips ?? i,
    },
    ...
}
```

`J` and `Q` are the module-level fallbacks in the same chunk:

```js
var J = 'Calibri',
  Q = 11;
```

That normalization is the only place `defaultSize` and `defaultFont` appear in the entire runtime build.
Counting literal occurrences across every `dist/*.mjs` and `dist/*.js` chunk:

| identifier            | occurrences | chunks | read anywhere?                |
| --------------------- | ----------- | ------ | ----------------------------- |
| `defaultSize`         | 4           | 2      | no, only the assignment above |
| `defaultFont`         | 4           | 2      | no, only the assignment above |
| `defaultTabStopTwips` | 28          | 12     | yes                           |
| `defaultFontSize`     | 12          | 6      | yes                           |
| `defaultFontFamily`   | 12          | 6      | yes                           |

The two chunks are `chunk-5UCYRCUV.mjs` and its CJS twin `chunk-LFJ2LWCA.js`, and both counts of 4 are the same two expressions per chunk (the property key plus the `s.defaultSize` / `s.defaultFont` read of the caller's value).
`defaultTabStopTwips` is the control: it is assigned on the very same line, and it is threaded through twelve chunks because it is genuinely consumed.
Property names survive minification, so an absent literal means an absent read.

The two internal callers of `toFlowBlocks` pass only `theme`, never the typography options:

```js
// dist/chunk-ZVLJ5D3Q.mjs
let t = c(e, { theme: r.theme ?? void 0 });

// dist/chunk-G2ADB4YD.mjs
a = c$2(r, { theme: o.theme ?? void 0 });
```

## Where the size actually comes from

Because nothing propagates a default, an unstyled run reaches measurement with `fontSize` undefined and picks up a hardcoded constant.
`dist/chunk-L7HQCG64.mjs`:

```js
var D = 11,
  X = 'Calibri',
  ...;

function q(e) {
  return {
    fontFamily: e.fontFamily ?? X,
    fontSize: e.fontSize ?? D,
    bold: e.bold,
    italic: e.italic,
    letterSpacing: e.letterSpacing,
  };
}
```

The same 11pt fallback is duplicated at five sites rather than centralized: `Q = 11` in `chunk-5UCYRCUV.mjs`, `r = 11` in `chunk-G2ADB4YD.mjs`, and `Ae = 11`, `Mt = 11`, `D = 11` in `chunk-L7HQCG64.mjs`.
The DOM painter hardcodes the matching value onto the page container instead of deriving it, in `chunk-G2ADB4YD.mjs`:

```js
e.style.fontFamily = a('Calibri').cssFallback;
e.style.fontSize = `${1056 / 72}px`; // 14.667px = 11pt
```

## Reproduction

`scripts/repro-default-typography-options.mjs` converts the same one-paragraph document three ways and compares the emitted blocks.
It needs an installed core build, because this repository's workspace has no real core to resolve against (see the last section):

```
node scripts/repro-default-typography-options.mjs --core <node_modules>/@eigenpal/docx-editor-core
```

Observed against `@eigenpal/docx-editor-core@1.9.0`:

| case | input                                                         | emitted run                                            |
| ---- | ------------------------------------------------------------- | ------------------------------------------------------ |
| A    | no options                                                    | `{ text: 'Hello' }`                                    |
| B    | `{ defaultSize: 24, defaultFont: 'Arial' }`                   | `{ text: 'Hello' }`                                    |
| C    | paragraph attr `defaultTextFormatting: { fontSize: 20, ... }` | `{ text: 'Hello', fontFamily: 'Arial', fontSize: 10 }` |

B is byte-identical to A: the documented options do not reach the output at all.
The script exits 1 on that comparison and 0 once the options are honored.

## Impact

- `defaultSize` and `defaultFont` are documented public options that do nothing. A caller has no way to discover this short of reading the shipped bundle.
- There is no supported replacement. Neither adapter's `DocxEditorProps` exposes a document default font size or family: `fontFamilies` only populates the toolbar dropdown and `fonts` only registers `@font-face` rules (`packages/react/src/components/DocxEditor.tsx`, `packages/vue/src/components/DocxEditor/types.ts`).
- Every size-less run renders at 11pt Calibri regardless of what the consuming application wants as its default.

## Workaround

Case C above is the only channel that works.
Injecting `w:sz` into `docDefaults.rPr` in `word/styles.xml` before loading the document sets `defaultTextFormatting` on each paragraph, and the value does reach the runs.
It is halved on the way in, per the half-point unit of `w:sz` (ECMA-376 §17.3.2.38): `fontSize: 20` yields 10pt.

This works, but it means changing a rendering default requires rewriting the document rather than configuring the editor.

## What a fix needs

The fix belongs in core, and the shape depends on intent:

- If the options are meant to work, `toFlowBlocks` should thread `defaultSize` / `defaultFont` into the emitted blocks the way `defaultTextFormatting` already is, and the two internal callers should forward whatever the host configured instead of passing `theme` alone.
- If they are not meant to work, they should be removed from `ToFlowBlocksOptions` so the type stops advertising them, and the 11pt fallback should be documented as fixed.

Either way the five duplicated `11` constants and the `1056 / 72` literal in the painter are worth collapsing to one source, since a default that is spelled out in six places cannot be changed reliably.

Separately, and deliberately out of scope here: ECMA-376 leaves the size unspecified when `w:sz` is absent, and Word's application default is 20 half-points (10pt), whereas core uses 11pt.
That is a fidelity question about which constant is correct, not about whether the option is wired up, and it should not be bundled into this fix.

## Why this report carries no code fix

The parser, layout engine, and layout bridge are not in this repository.
`packages/core` is `@docx-editor.dev/core-contract`, described in its own `package.json` as a "Contract-only declaration of the @docx-editor.dev/core public API. Not the implementation; never published", and its `src/` holds six declaration files with every function throwing `contract-only stub: no implementation`.

Concretely, in this repository:

- `toFlowBlocks`, `ToFlowBlocksOptions`, and the string `layout-bridge` do not appear anywhere in `packages/`, `e2e/`, `examples/`, `openspec/`, or `docs/` outside this report and its reproduction script. Not even the Vue adapter, which the deep import's own doc comment names as a consumer, imports it here.
- `packages/core/src/` contains no `flow-model/` directory, so the paths that `CONTRIBUTING.md` and the `openspec/changes/` proposals cite for layout work (for example `packages/core/src/flow-model/footnoteLayout.ts`) resolve only against the private implementation.
- `bun.lock` pins `@docx-editor.dev/core` to the published `0.0.1`, whose npm tarball unpacks to 770 bytes, so `bun test` cannot exercise a real core and a regression test for this cannot live in the repository's suite. That is why the reproduction is a standalone script pointed at an installed build.

So this is a report plus a reproduction.
The behavior change has to happen in the core implementation.
