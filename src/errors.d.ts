/**
 * gridxcel/errors - the one error hierarchy.
 *
 * Every failure the library raises is an instance of \`GridxcelError\`, so a
 * single \`instanceof\` check covers the whole surface. Each carries a stable
 * machine-readable \`code\` that is safe to branch on; the message is for humans
 * and may change.
 */

/** What the caller did wrong or what the input violated. */
export type GridxcelErrorCode = string;

export declare class GridxcelError extends Error {
  constructor(message: string, code?: GridxcelErrorCode, details?: Record<string, unknown>);
  /** Stable, machine-readable discriminator. Safe to compare. */
  code: GridxcelErrorCode;
  /** Structured context: the part, the offset, the offending value. */
  details?: Record<string, unknown>;
}

/** Container-level failure: the bytes are not a usable ZIP, or a part is corrupt. */
export declare class BadZipError extends GridxcelError {}
/** The bytes are a ZIP but not a workbook, or the workbook is malformed. */
export declare class FormatError extends GridxcelError {}
/** Well-formed XML that violates the OOXML schema. */
export declare class SchemaError extends GridxcelError {}
/** A feature the library deliberately does not implement was requested. */
export declare class NotImplementedError extends GridxcelError {}
/** A warning raised as an error because \`strict: true\` was set. Carries the warning. */
export declare class StrictWarningError extends GridxcelError {
  warning?: { code: string; message: string; [key: string]: unknown };
}

/**
 * A host capability the library needs is missing or unusable.
 *
 * Always a GridxcelError (so catching by type keeps working) and always code
 * NOT_SUPPORTED (so you can branch on the code alone). Raised outside Node by
 * the synchronous API and by the path-based helpers, because a browser has no
 * synchronous inflate and no filesystem. Every one of those has an async twin.
 */
export declare class PlatformError extends GridxcelError {}
