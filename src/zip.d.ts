/** gridxcel/zip - the ZIP container layer, for callers that need it directly. */

/**
 * The streaming subset the library relies on.
 *
 * Declared structurally rather than as \`NodeJS.ReadableStream\` so the published
 * types do not require @types/node. A browser consumer, or a Node consumer who
 * has not installed the Node types, can use every API without a stub - and
 * Node's \`stream.Readable\` satisfies this interface as it is.
 */
export interface ByteStream extends AsyncIterable<Uint8Array> {
  pause?(): unknown;
  resume?(): unknown;
  destroy?(error?: Error): unknown;
  on?(event: string, listener: (...args: unknown[]) => void): unknown;
}

import type { BadZipError } from "./errors.js";

export declare const SIG_LOCAL: number;
export declare const SIG_CENTRAL: number;
export declare const SIG_EOCD: number;
export declare const SIG_EOCD64: number;
export declare const SIG_EOCD64_LOC: number;
/** 0 = stored, 8 = deflate. */
export declare const METHOD_STORE: number;
export declare const METHOD_DEFLATE: number;

export interface ZipArchiveOptions {
  /**
   * Verify every decompressed part against the CRC-32 in the central
   * directory. Default true. A mismatch throws ZipError/CRC_MISMATCH instead
   * of returning corrupted bytes. Set false only if you verify integrity
   * yourself; the cost is roughly 1.2 GB/s of payload.
   */
  verifyCrc?: boolean;
}

export declare class ZipEntry {
  name: string;
  size: number;
  compressedSize: number;
  method: number;
  crc: number;
  /** Offset of the local file header within the archive. */
  localOffset: number;
  archive: ZipArchive;
  /** Decompressed bytes, cached after the first call. Throws ZipError/CRC_MISMATCH. */
  bytes(): Uint8Array;
  /** Decompressed content decoded as UTF-8. */
  text(): string;
  /**
   * Decompressed bytes, asynchronously.
   *
   * The twin of \`bytes()\` for platforms without synchronous inflate - a browser
   * has DecompressionStream and nothing synchronous. Prefer this in any code
   * that must also run outside Node; \`bytes()\` throws PlatformError
   * [NOT_SUPPORTED] there rather than returning something wrong.
   */
  bytesAsync(): Promise<Uint8Array>;
  /** Decompressed content decoded as UTF-8, asynchronously. Twin of \`text()\`. */
  textAsync(): Promise<string>;
  /** Live inflate stream, verified incrementally. */
  stream(): ByteStream;
}

export declare class ZipArchive {
  /**
   * @param input complete archive bytes
   */
  constructor(input: Uint8Array | ArrayBuffer, opts?: ZipArchiveOptions);
  /** Read and parse a file from disk. */
  static open(path: string): Promise<ZipArchive>;
  entries: ZipEntry[];
  byName: Map<string, ZipEntry>;
  comment: string;
  size: number;
  verifyCrc: boolean;
  get(name: string): ZipEntry | undefined;
  has(name: string): boolean;
  read(name: string): Uint8Array | undefined;
  readText(name: string): string | undefined;
  names(): string[];
  [Symbol.iterator](): Iterator<ZipEntry>;
}

export interface ZipWriterOptions {
  /** Deflate level 0-9. 0 stores every part. Default 6. */
  compressionLevel?: number;
  /** Fixed DOS timestamps for byte-identical output across runs. Default false. */
  deterministic?: boolean;
}

export interface ZipAddOptions {
  compress?: boolean;
  level?: number;
  date?: Date;
}

export declare class ZipWriter {
  constructor(opts?: ZipWriterOptions);
  add(name: string, data: Uint8Array | string, opts?: ZipAddOptions): this;
  build(): Uint8Array;
  /** Alias of \`build()\`, for callers who expect a Buffer-shaped name. Returns bytes. */
  toBuffer(): Uint8Array;
}

export declare function createZip(
  files: Iterable<[string, Uint8Array | string]> | Record<string, Uint8Array | string>,
  opts?: ZipWriterOptions,
): Uint8Array;

export declare class ZipError extends BadZipError {
  code: string;
  offset?: number;
}
