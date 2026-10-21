# Contributing

gridxcel is a small library with a large verification suite. This file covers
what you need to build, test and change it, and the conventions the code already
follows. Participation is covered by the [Code of Conduct](CODE_OF_CONDUCT.md).

## Set up

Clone the repository and open it. That is the whole setup:

- Node.js 18 or newer.
- No dependencies to install. Both dependency lists in package.json are empty,
  and that is a design decision rather than a detail: adding one is worth an
  issue before it is worth a pull request.
- No build step. `src/*.js` is what ships and what the tests run; there is no
  `dist/`, no transpile and no bundler.
- ESM only: `import` and `export`, no CommonJS build, no dual package.
- The suite needs no network access and no install step. A change that makes a
  test fetch a URL, download a fixture, or require a compile breaks that, so it
  will not be merged.

## Run the checks

| Command | What it does |
| --- | --- |
| `npm test` | The `node:test` suite: unit, integration, interop and robustness. As last measured, 324 tests, 320 pass, 0 fail, 4 skipped; the skips are fixtures the Python oracle refuses to open. |
| `npm run check:docs` | Documentation checker: every backticked symbol must exist in `src/`, every local link and anchor must resolve, and every `doc-example` fence is executed. |
| `npm run interop` | The oracle harness: gridxcel writes and Python openpyxl verifies, then gridxcel reads Python-written fixtures and diffs them. Exit 0 pass, 1 mismatch, 2 nothing proven. |
| `npm run check:exports` | Every public export appears in a document, and every documented symbol exists. The entry points are derived from `package.json`. |
| `npm run check:clone` | Runs the documentation checks against a copy of the tracked files only, with no bench directory and no installed dependencies. Catches a doc example that depends on something not in the repository - which is invisible on a developer machine by definition. |
| `npm run check:fuzz` | Five seconds of generated input across ten parsers, with a fixed seed. `npm run fuzz:long` runs it for two minutes. |
| `npm run check:release` | The pre-flight for a publish. See [Releasing](#releasing). |
| `npm run bench:scaling` | The scaling and head-to-head numbers behind [docs/PERFORMANCE.md](docs/PERFORMANCE.md). |
| `node scripts/check-crc.mjs` | CRC regression check. Exit 1 means corruption was rejected, which is the good outcome; exit 0 means the gap is back. |

## Documented examples are executed

A fence whose info string contains `js doc-example` is not decoration. The
checker extracts it, writes it to `bench/doc-examples/` and runs it with Node:

```js doc-example
import { read, write, sheetToJSON } from "gridxcel";
import assert from "node:assert";

const bytes = await write([["name", "score"], ["ada", 42]], { bookType: "xlsx" });
const book = await read(bytes);
const rows = await sheetToJSON(book.sheets[0], { header: 1 });
assert.deepStrictEqual(rows, [["name", "score"], ["ada", 42]]);
console.log("round trip ok:", rows.length, "rows");
```

Two things follow from that:

- **Adding one means it must pass.** A broken example fails `npm run check:docs`
  and the workflow that runs it, so it cannot rot quietly.
- **It runs from `bench/doc-examples/`**, so a relative path inside an example
  resolves from there; the README example reaches back to the fixtures with
  `../../test/fixtures/golden.xlsx`. The last line the example prints is echoed by
  the checker, so a `console.log` is a cheap way to show what the code did.

The same checker also reads every backticked identifier in the documents it
covers and requires it to be exported by `src/` or to appear in the source. A
name that is not API - an environment variable, another tool's option, a CLI
verb - goes in `scripts/docs-allow.json` or is written without a code span.

## Releasing

A release is the one irreversible action here: an npm version can be deprecated
but never withdrawn. Two things make it safe.

**`npm run check:release`** runs the pre-flight locally. It refuses when the
version is not valid semver, when it is already on the registry, when the git tag
and the manifest disagree, when the tree is dirty, when a dependency has crept
in, when an exports target does not exist, or when the packed tarball does not
install and resolve every entry point. That last one is not redundant with the
test suite: it is the only check that looks at what actually ships.

**`.github/workflows/release.yml`** publishes from CI, triggered by a `v*` tag.
Publishing from CI is not convenience - it is the only way to get npm provenance,
because the attestation needs an OIDC token that only a supported provider can
issue. A laptop publish is an unattested publish.

To cut a release:

```sh
npm run check:release          # must print "ready to publish"
git tag v0.1.0
git push origin v0.1.0         # the workflow publishes
```

The workflow needs an NPM_TOKEN repository secret. It runs the full gate plus
the Python oracle on a clean checkout before publishing, and records the tarball
URL and integrity digest in the run summary. It also refuses to publish when the
tag and the manifest disagree, because that would put a version on the registry
that no tag in the repository points at.

Use gh workflow run release.yml with the dry-run input for a rehearsal: it builds and
verifies without publishing.

### Choosing the version

An npm version is immutable. Publishing 0.1.0 freezes this API surface
permanently - it can be deprecated, never corrected. Two honest choices:

| Version | npm dist-tag | When |
| --- | --- | --- |
| `0.1.0` | the default tag | The surface is settled and you will support it. |
| `0.1.0-rc.1` | the next tag | You want the artifact in front of people while keeping room to change. npm assigns that tag to a prerelease automatically. |

A prerelease reserves nothing permanently and is not offered to a plain install
by default; users opt in by naming the tag. Given that this
project's API grew three ways during the 0.1.0 build (`createCSVStream`,
`gridxcel/formats`, and the corrected type declarations), a prerelease is the
lower-regret choice unless the surface is genuinely final.

## Fixtures and the Python oracle

- Fixtures live in `test/fixtures/`. Most are written by Python rather than by
  hand: `test/make_fixtures.py` (openpyxl and XlsxWriter), `test/make_golden.py` and
  `test/make_reader_fixtures.py`. Commit the generator next to the fixture so the
  file can be rebuilt instead of reverse-engineered.
- Add a fixture when the existing ones cannot express the case and the case
  matters - a malformed package, a container quirk, a number-format edge - and
  keep it as small as it can be while still showing the behaviour.
- A test needs the Python oracle when the claim is "another implementation
  agrees with us" rather than "we agree with ourselves": the interop reads and
  writes, the golden manifests, number-format classification. Those tests reach
  Python through `scripts/oracle-client.mjs`; set `GRIDXCEL_PYTHON` to point at a
  different interpreter.
- The oracle is openpyxl and XlsxWriter, never another gridxcel module. If both
  sides of a check came out of `src/`, the check proves nothing.
- A suite that cannot find the oracle, or a module it needs, skips with a
  `NOT YET AVAILABLE` reason instead of failing, so a partial checkout still runs
  green. If you add an oracle test, check that it ran rather than skipped.

## Conventions

- **Errors are typed.** Throw a subclass of `GridxcelError` carrying a stable
  machine-readable code drawn from the existing set; callers branch on the code,
  so treat a published code as API.
- **Warnings are data.** The reader collects problems and routes them through
  `onWarning`; `strict` promotes them to a `StrictWarningError`. Do not throw
  from deep inside a parser when a warning would carry the same information.
- **Sync and async are one implementation.** The second path is a thin adapter,
  not a copy. Two copies of a helper is how the open duplication defects
  happened.
- **Comments explain why.** A comment that repeats the line beneath it is noise;
  the ones worth writing record a decision - why the scanner is hand-rolled, why
  this problem is a warning and not an error.
- **British prose in comments**, consistently within a file: `normalisation`,
  `behaviour`, `colour`.
- **Declarations are hand-written** and are part of the API, so update the
  matching `.d.ts` when the public surface changes.
- **Dates use the local wall clock.** Never compare a parsed date with
  `toISOString()`; the suite is checked under several time zones for that reason.

## Pull requests

- One topic per pull request, with a narrow diff.
- Run `npm test` and `npm run check:docs` before pushing, and say in the
  description what you ran and what it printed.
- Add a `CHANGELOG.md` entry under Unreleased when behaviour changes, and say
  plainly in the description anything you could not verify.
- Do not add a dependency, a build step or a network call in a test without
  saying so in the description. Those are load-bearing properties of this
  project, not preferences.
- A performance claim needs the command, the input file and the machine. A
  number without a reproduction cannot be reviewed.
- Security reports do not belong in a pull request or an issue: see
  [SECURITY.md](SECURITY.md).
