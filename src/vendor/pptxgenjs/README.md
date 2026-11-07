# Vendored PptxGenJS 4.0.1

Unmodified copy of `dist/pptxgen.cjs.js` (-> `index.js`) and `types/index.d.ts`
from `pptxgenjs@4.0.1` (MIT, see `LICENSE`), https://github.com/gitbrent/PptxGenJS

Why: the npm package declares `image-size ^1.2.1`, which has two unpatched
advisories (CVE-2025-71329, CVE-2025-71330). The published bundle never loads
`image-size` (upstream issues #1473, #1474, #1529). Vendoring removes that
declared dependency from `pptx-templates`. Only `jszip` is needed at runtime.

No code changes were made. To update: copy the same two files from a newer
`pptxgenjs` release and re-run the test suite.
