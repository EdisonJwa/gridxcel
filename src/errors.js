/** Error taxonomy: every failure carries a stable machine-readable code. */

export class GridxcelError extends Error {
  constructor(message, code = "GRIDXCEL_ERROR", details) {
    super(message);
    this.name = "GridxcelError";
    this.code = code;
    if (details !== undefined) this.details = details;
  }
}

/** Bad zip container. */
export class BadZipError extends GridxcelError {
  constructor(message, code = "ZIP_ERROR", details) {
    super(message, code, details);
    this.name = "BadZipError";
  }
}

/** Part is not a valid xlsx/OPC package. */
export class FormatError extends GridxcelError {
  constructor(message, code = "BAD_FORMAT", details) {
    super(message, code, details);
    this.name = "FormatError";
  }
}

/** Well-formed XML that violates the OOXML schema. */
export class SchemaError extends GridxcelError {
  constructor(message, code = "BAD_SCHEMA", details) {
    super(message, code, details);
    this.name = "SchemaError";
  }
}

/** A feature the caller asked for is not implemented. */
export class NotImplementedError extends GridxcelError {
  constructor(message, details) {
    super(message, "NOT_IMPLEMENTED", details);
    this.name = "NotImplementedError";
  }
}

/** Thrown by strict mode on the first malformed part. */
export class StrictWarningError extends GridxcelError {
  constructor(warning) {
    super(warning.message, warning.code, warning);
    this.name = "StrictWarningError";
    this.warning = warning;
  }
}

/**
 * A host capability the library needs is missing or unusable.
 *
 * Always a GridxcelError, so catching by type keeps working, and always code
 * NOT_SUPPORTED, so a caller can branch on the code alone. Raised outside Node
 * by the synchronous API (readSync, writeSync, ZipArchive.bytes()) and by the
 * path-based helpers, because a browser has no synchronous inflate and no
 * filesystem. Every one of those has an async twin.
 */
export class PlatformError extends GridxcelError {
  constructor(message, details) {
    super(message, "NOT_SUPPORTED", details);
    this.name = "PlatformError";
  }
}
