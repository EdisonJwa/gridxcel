# Using gridxcel in a browser

gridxcel runs in a browser with no build step and no polyfill. The same
`src/*.js` files that Node loads are what a bundler loads; there is no separate
browser bundle to keep in sync, and no dependency is added to get there.

Everything on this page is executed by `integration/browser-bundle.test.mjs`,
`integration/vite.test.mjs` and `integration/react.test.mjs`. If a claim here
stops being true, those fail.

---

## What works, and what does not

| | Browser |
| --- | --- |
| `read`, `eachRow`, `sheetToJSON`, `sheetToCSV`, `write`, `WorkbookWriter` | Yes |
| Reading from a `File`, `Blob` or `ArrayBuffer` | Yes |
| Streaming a large sheet without materialising it | Yes |
| `readSync`, `writeSync`, `ZipArchive.bytes()` | **No** - throws `PlatformError` `[NOT_SUPPORTED]` |
| `readFile`, `writeFile` and other path-based helpers | **No** - throws `PlatformError` `[NOT_SUPPORTED]` |

The synchronous API needs a synchronous inflate, and a browser only has
`DecompressionStream`, which is asynchronous. So the synchronous entry points
are Node-only, and outside Node they fail with a typed error naming the async
alternative rather than a `ReferenceError`. Every synchronous entry point has an
async twin, so nothing is actually lost.

Path-based helpers are Node-only for the obvious reason: a browser has no
filesystem. Use a `File` or a `fetch` response instead.

## Reading a file the user picked

```js
import { read, sheetToJSON } from "gridxcel";

async function load(file) {
  // file is from an <input type="file"> or a drag-and-drop event
  const bytes = new Uint8Array(await file.arrayBuffer());
  const book = await read(bytes);
  return sheetToJSON(book.sheets[0]); // array of objects
}
```

`read` accepts a `Uint8Array`, an `ArrayBuffer`, or anything async-iterable of
bytes, so a `fetch` body streams straight in without being buffered first:

```js
const response = await fetch("/report.xlsx");
const book = await read(response.body);
```

## Streaming a sheet too large to hold

`eachRow` reads row by row, so a 200,000-row sheet never becomes a 200,000-row
array:

```js
import { eachRow } from "gridxcel";

for await (const row of eachRow(bytes)) {
  if (row.cells.length) render(row);
}
```

## Building a workbook in the browser

```js
import { write } from "gridxcel";

const bytes = await write([["name", "score"], ["ada", 42]], { bookType: "xlsx" });
// bytes is a Uint8Array: hand it to a Blob, or POST it
const url = URL.createObjectURL(new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
```

Note that a browser has no synchronous deflate, so the writer **stores** parts
instead of compressing them. The result is a valid `.xlsx` that Excel, openpyxl
and gridxcel all read; it is simply larger. This is a deliberate trade: a bigger
file beats a failed write.

---

## Bundlers

Nothing to configure. There is no `node:` import in the browser build at all,
so there is nothing to shim, stub or externalise.

```js
import { read, write } from "gridxcel";
```

Verified with **esbuild**, **Vite** and **rolldown** (each built, then executed
in a sandbox with no `Buffer`, `process`, `require` or module loader) and with
**webpack 5** in a real browser.

Best of all, `integration/real-browser.test.mjs` launches the Chrome or Edge
installed on the machine through puppeteer-core, loads a real page over
`file://`, and runs the library in the page's own JavaScript context - a real
engine, real streams, real `atob`, with `Buffer`, `process` and `require` all
`undefined`. It reads a deflated workbook, converts it, streams 205 rows, writes
a workbook and reads it back, and confirms `readSync` throws `PlatformError`
`[NOT_SUPPORTED]` rather than a `ReferenceError`.

### Tree-shaking

Importing one function does not pull in the library. Measured on the current
source: importing `sheetToCSV` alone produces a **21 kB** minified bundle
against **251 kB** for the whole surface.

### The `browser` field

`package.json` maps one module:

```json
"browser": { "./src/platform.js": "./src/platform.browser.js" }
```

`src/platform.js` is the Node implementation and imports `node:zlib`,
`node:stream`, `node:fs` and `node:url`. `src/platform.browser.js` is the Web
implementation: `CompressionStream`/`DecompressionStream`, and no `node:`
reference anywhere. Bundlers that honour the standard `browser` field pick the
right one, and four are exercised by this repository's own tests: **esbuild**,
**Vite**, **rolldown** (the Rust bundler inside Vite 8) and **webpack 5**. The
first three build and run in a bare sandbox; webpack builds and runs in a real
browser, because webpack turns the library's lazy peer imports into a separate
chunk that it loads on demand - correct behaviour, but one that needs a browser
to exercise. Classic Rollup honours the same documented field and is not run
here, so treat it as supported-by-standard rather than verified-by-test. Both files export an identical 19-name
surface, asserted by `test/platform.test.mjs`, and so is the mapping itself - the
swap cannot quietly rot.

This is also why the split exists rather than a runtime check: Node 18 has no
`process.getBuiltinModule`, so a single file that sniffed the platform at runtime
would have had to break either the browser build or Node 18's synchronous API.

### `<script>` tags

The package is ESM, so use `<script type="module">`. There is no top-level
`await` anywhere in `src/`, which is what makes an IIFE build possible at all;
an earlier design had one in `gridxcel/compat` and esbuild refused the bundle
outright. If you need a classic non-module script, bundle it yourself - the
output of that build is tested.

---

## React, Preact and anything else

The core is plain data in, plain data out. There is no framework adapter because
none is needed: a `Workbook` is a plain object, and it survives
`structuredClone`, so it can cross a worker boundary or sit in React state.

```jsx
// A React server component reading a workbook and rendering a table.
import { read, sheetToJSON } from "gridxcel";

export async function SheetTable({ bytes }) {
  const book = await read(bytes);
  const rows = await sheetToJSON(book.sheets[0], { header: 1 });
  const [head, ...body] = rows;
  return (
    <table>
      <thead><tr>{head.map((c, i) => <th key={i}>{String(c)}</th>)}</tr></thead>
      <tbody>
        {body.map((r, i) => <tr key={i}>{r.map((c, j) => <td key={j}>{String(c)}</td>)}</tr>)}
      </tbody>
    </table>
  );
}
```

React 19 server rendering, Preact's renderer and `structuredClone` are all
exercised by `integration/react.test.mjs`.

---

## TypeScript

The package ships hand-written `.d.ts` files and needs no `@types/node`. The
stream type is declared structurally as `ByteStream` rather than as
`NodeJS.ReadableStream`, precisely so a browser consumer does not have to
install Node's types to use a library that does not need Node.

`integration/types.test.mjs` compiles a real consumer against the package with
`skipLibCheck: false` and no `@types/node`, which is what a browser or bundler
consumer has. That test is why the declarations are trustworthy: before it
existed, `src/index.d.ts` shipped a syntax error and three exports had no
declaration at all.

---

## Bundle size

| | Minified |
| --- | --- |
| `import { sheetToCSV } from "gridxcel"` | 21 kB |
| Every export | 251 kB |

There is no runtime dependency, so there is nothing else in the graph.

---

## Running the browser tests

```sh
cd integration
npm install
node run-all.mjs
```

The harness has its own dev tooling (esbuild, Vite, React, Preact, TypeScript)
and is deliberately **not** published: `package.json` `files` excludes it, and
`npm pack` confirms it. The package stays zero-dependency; only the tests around
it have tooling.
