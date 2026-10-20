# Performance

What the parser costs, how that was measured, and the one defect that dominated
everything else. Reproduce any number here with:

```
node --expose-gc scripts/bench-scaling.mjs          # scaling + head to head
node --expose-gc scripts/bench.mjs --rows=200000    # throughput + peak memory
```

## The finding: the scanner was quadratic in the buffer size

The reader was six times slower than its own streaming path on identical input -
3,558 ms against 540 ms for the same 400,008 cells. Both paths run the same
`RowScanner`, so the difference had to be in how it was fed, and it was:
the synchronous path feeds 256 KiB slices, the streaming path receives whatever
zlib emits, around 16 KiB.

Parsing the same document at a range of feed sizes exposed the shape of the
problem. These are the original numbers, median of three:

| Feed size | Time | vs 16 KiB |
| --- | --- | --- |
| 16 KiB | 724 ms | 1.0x |
| 64 KiB | 1,549 ms | 2.1x |
| 256 KiB | 4,296 ms | 5.9x |
| 1 MiB | 18,730 ms | 25.9x |

A linear scanner is flat across those sizes. That curve is quadratic.

### Root cause

`findTagEnd` in `sheet-reader.js` had to find the end of a tag while
respecting the rule that a `>` inside a quoted attribute value does not end
the tag. It looked for the next `>`, then asked whether a quote appeared before it:

```js
const gt = s.indexOf(">", lt + 1);
const dq = s.indexOf('"', lt + 1);
if (dq < 0 || dq > gt) {
  const sq = s.indexOf("'", lt + 1);   // <-- unbounded
  if (sq < 0 || sq > gt) return gt;
}
```

Neither `indexOf` call is bounded by the tag. They search to the end of the
buffer. For a tag that has no attributes the very next character is `>`, so the
`>` is found immediately, the double quote is usually far ahead, and the
apostrophe search then scans the **entire remaining buffer** looking for a
character that, in this document, does not exist at all.

Measured on `bench/fair-50000x8.xlsx`:

* **1,650,050 attribute-less tags** (`<v>`, `<t>`, `<is>` and every closing tag)
* **0 apostrophes** in all 19,766,218 characters

Every one of those 1.65 million tags scanned to the end of its buffer and found
nothing, which is why the cost scaled with the square of the buffer size.

### The fix

Walk the tag once, tracking quote state, and stop at the first `>` that is
outside a quoted value. The scan is bounded by the tag itself:

```js
function findTagEnd(s, lt) {
  let quote = 0;
  const n = s.length;
  for (let i = lt + 1; i < n; i++) {
    const c = s.charCodeAt(i);
    if (quote !== 0) { if (c === quote) quote = 0; continue; }
    if (c === 34) { quote = 34; continue; }
    if (c === 39) { quote = 39; continue; }
    if (c === 62) return i;
  }
  return -1;
}
```

This is the same algorithm `xml.js` already used, so the two scanners now agree.
It is equivalent to the old logic - the fast path returned the first `>` when
no quote preceded it, which is exactly what the quote-aware walk returns - and
`test/scaling.test.mjs` pins both the equivalence and the scaling.

### Result

| Feed size | Before | After | Speedup |
| --- | --- | --- | --- |
| 16 KiB | 724 ms | 289 ms | 2.5x |
| 64 KiB | 1,549 ms | 283 ms | 5.5x |
| 256 KiB | 4,296 ms | 285 ms | 15.1x |
| 1 MiB | 18,730 ms | 307 ms | 61.0x |

The curve is flat now, which is the property worth protecting. End to end on the
same file, the fix removed the gap between the two readers entirely:

| Path | Before the fix | After the fix |
| --- | --- | --- |
| `readWorkbookSync` | 3,558 ms | 519 ms |
| `readWorkbook` (async) | 540 ms | 520 ms |
| sync / async ratio | **6.59x slower** | **1.00x** |

The synchronous reader no longer has a reason to exist as the slow path, and the
gap was never architectural: the two paths share one scanner and differed only in
how the scanner was fed.

Absolute times on this machine move by a factor of two between runs, so the ratio
is the durable result. Both columns were measured with the same harness on the
same input, before and after the change.

## Head to head against SheetJS CE 0.20.3

Same bytes, same process, same machine; both libraries read all 400,008 cells.
```sh
node --expose-gc scripts/bench-scaling.mjs --passes=7
```

Seven passes, each the median of three runs; the time shown is the median across
passes, and the speedup is reported as the range across those same passes.

| Operation | Median time | Versus SheetJS |
| --- | --- | --- |
| SheetJS `read` | 961 ms | - |
| gridxcel `readSync` | 519 ms | median 1.85x, range 1.57x - 2.37x |
| gridxcel `read` (async) | 520 ms | median 1.91x, range 1.42x - 2.62x |
| gridxcel `eachRow` (streaming) | 511 ms | median 1.88x, range 1.56x - 2.68x |
| SheetJS read + write | 1,427 ms | - |
| gridxcel `writeWorkbookSync` | 348 ms | write only |

**The headline is the range: roughly 1.4x to 2.7x faster, median about 1.9x.**

### A correction, and why the range matters

Earlier versions of this document and of `COMPARISON.md` quoted
2.74x/3.11x/3.06x, and before that 4.01x/4.38x/4.66x. Those were real
measurements, but they were **single passes**, and a single pass on a shared
machine is a sample, not a result: process-level effects (GC timing, CPU
frequency, whatever else is running) move the two libraries independently.
Repeating the identical command produced speedups from 1.42x to 4.66x.

Publishing the best sample as "the" speedup is the most common way a benchmark
stops being trustworthy, so this document now publishes the distribution and the
benchmark reports it directly. If you re-run the command, expect a number inside
the range rather than the median.

The direction is unambiguous and has held in every run: gridxcel's reader is
faster than SheetJS CE 0.20.3 on this input, and its streaming path costs the
same as a full parse while holding a row instead of a workbook.

A second correction: an earlier benchmark measured gridxcel's synchronous read as
3.7x **slower** than SheetJS. That was the quadratic defect described above, not
a property of the design.

## Memory

Measured by `scripts/bench.mjs` on a 200,000 x 10 workbook; see
`bench/bench-200000x10.json`. Peak heap is sampled on a 5 ms interval and the
retained figure is taken after a forced collection, so both the transient cost
and what stays alive are visible.

| Stage | ms | rows/s | peak heap | retained |
| --- | --- | --- | --- | --- |
| `eachRow` (streaming) | 2,587 | 77,309 | 13.0 MiB | 5.9 MiB |
| `readWorkbook` (full parse) | 2,937 | 68,086 | 224.9 MiB | 164.9 MiB |
| `sheetToCSV` | 414 | 482,905 | 259.6 MiB | 184.2 MiB |
| `WorkbookWriter` (streaming write) | 3,446 | 58,032 | 375.9 MiB | 5.8 MiB |
| `writeWorkbookSync` | 2,769 | 72,239 | 534.1 MiB | 164.1 MiB |

Streaming holds 5.9 MiB retained against 164.9 MiB for a full parse, at
comparable throughput. That is the reason `eachRow` exists.

## What did not matter

Recorded so the same ground is not covered twice:

* **Inflate.** 24 ms of a 3,000 ms parse. zlib is not the problem and no amount
  of tuning it would have helped.
* **UTF-8 decoding.** 8 ms to decode all 18.85 MiB in one call, 24 ms streamed in
  256 KiB slices.
* **Chunk size, after the fix.** Flat from 16 KiB to a single 19 MiB chunk. The
  earlier sensitivity was the quadratic, not the chunking strategy.
* **Tag parsing.** `parseTag` costs 12 ms per million attribute-less tags and
  76 ms per million with attributes. Even eliminating it entirely would be
  rounding error next to the 3,000 ms the quadratic was consuming.

That last point is worth stating plainly, because it was the wrong first guess:
the profile looked like "lots of small allocations in the tag path", and
optimising that would have bought a few percent while missing the quadratic,
which cost up to 61x at large chunk sizes.

## Open opportunities

1. **`readWorkbookSync` still decodes the whole part into a string before
   scanning.** Decoding 18.85 MiB costs 8 ms in one call and 24 ms streamed, so it
   is a small share of a ~520 ms parse, but a scanner that worked directly on
   bytes would remove one full copy of the document from the pipeline.
2. **Cell object construction is now the largest remaining cost.** A full parse
   retains 164.9 MiB for 2 million cells, about 82 bytes per cell. The streaming
   path avoids it; a parse that builds cells lazily would not.
3. **`verifyCrc` costs about 1.6%** (3,724.7 ms against 3,785.9 ms on the older
   measurement harness). Not worth optimising, and worth keeping on.
4. **Writing is now the slower half** of a read-modify-write cycle. `writeWorkbookSync`
   has not had the same treatment the reader just received.

## Regression protection

`test/scaling.test.mjs` asserts that the 1 MiB / 16 KiB time ratio stays under
4. With the quadratic reintroduced the test reports 20.7x and fails with the
measured numbers in the message. The threshold is deliberately loose so the test
constrains complexity rather than machine speed.
