/** Values (and functions) the expressions in template tags can reference. */
export type TemplateData = Record<string, unknown>;

export type RenderTemplateOptions = {
  /** Default: '{{' */
  openingTag?: string;
  /** Default: '}}' */
  closingTag?: string;
  /**
   * docx-templates spelling of the delimiters: one string for both sides
   * (`'+++'`) or `[opening, closing]`. `openingTag`/`closingTag` win.
   */
  cmdDelimiter?: string | [string, string];
  /** Throw on unknown identifiers instead of treating them as undefined. */
  strict?: boolean;
  /**
   * Extra names (helpers, constants) for expressions, looked up after `data`,
   * so `data` can shadow them.
   */
  additionalJsContext?: Record<string, unknown>;
  /** Throw when an insertion evaluates to null or undefined. Default: false */
  rejectNullish?: boolean;
  /**
   * Called when a command fails to evaluate; its return value is used as the
   * command's result (inserted text, IF condition, FOR list, …).
   */
  errorHandler?: (error: Error, command: string) => unknown;
  /**
   * `false`: keep going after an evaluation error and throw one error listing
   * all failed commands at the end. Default: true (throw on the first error).
   */
  failFast?: boolean;
  /** Turn `\n` in inserted values into line breaks. Default: true */
  processLineBreaks?: boolean;
};

export type TemplateCommandType =
  | 'INS'
  | 'IF'
  | 'ELSE-IF'
  | 'ELSE'
  | 'END-IF'
  | 'FOR'
  | 'END-FOR'
  | 'EXEC'
  | 'ALIAS'
  | 'HTML'
  | 'LINK'
  | 'IMAGE';

/** Value of an `{{IMAGE …}}` tag (docx-templates compatible). */
export type TemplateImage = {
  /** Image contents: Buffer, ArrayBuffer, base64 string or data URL. */
  data: Buffer | ArrayBuffer | Uint8Array | string;
  /** '.png', '.jpg', '.jpeg' or '.gif'; optional for data URLs. */
  extension?: string;
  /** Width in cm. Omit both: fit into the shape, keeping the aspect ratio. */
  width?: number;
  /** Height in cm. Only one of width/height: the other keeps the aspect. */
  height?: number;
  /** Alternative text. */
  alt?: string;
  /** Rotation in degrees, clockwise. */
  rotation?: number;
};

export type TemplateCommand = {
  /** 1-based slide number (file name based: ppt/slides/slide<N>.xml). */
  slide: number;
  /** Name of the shape that holds the command. */
  shape?: string;
  /** The tag as written, e.g. `{{IF user.vip}}`. */
  raw: string;
  type: TemplateCommandType;
  /** The command without keyword, e.g. `user.vip`. */
  code: string;
};

export type ListTemplateCommandsOptions = Pick<
  RenderTemplateOptions,
  'openingTag' | 'closingTag' | 'cmdDelimiter'
>;
