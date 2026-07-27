---
'@docx-editor.dev/react': patch
---

Fix table row insertion severing vertical merges. `addRow`/`createEmptyRow` now emit a `vMerge: 'continue'` cell for any column whose merge spans the insertion boundary, so inserting a row inside a merged region grows the merge instead of dropping a phantom cell and shifting the cells below out of the grid.
