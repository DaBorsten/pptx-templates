import {
  CornerRadius,
  ImageBorder,
  ShapeAdjustments,
  ShapePreset,
} from '../shape-geometry/types';

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
  /**
   * Replace typographic quotes from AutoCorrect with straight ones before a
   * command is parsed, as docx-templates does: `“ ” „` → `"`, `‘ ’ ‚` → `'`.
   * This applies to the whole command, also to quotes inside a string
   * (`'Ada’s'` then ends after `Ada`). Default: false
   */
  fixSmartQuotes?: boolean;
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
  /**
   * Shape the picture is cropped to ("Crop to Shape"): any DrawingML preset
   * geometry, e.g. 'roundRect', 'ellipse', 'heart'. Defaults to 'rect', or to
   * 'roundRect' when `cornerRadius` is given.
   */
  shape?: ShapePreset;
  /**
   * Corner radius on the shorter side of the picture: a number in cm, or a
   * string like '10%', '12px', '4pt'. '50%' is the maximum. Only for the
   * rounded shapes ('roundRect', 'round1Rect', 'round2SameRect',
   * 'round2DiagRect', 'snipRoundRect').
   */
  cornerRadius?: CornerRadius;
  /** Raw adjust values of the shape, e.g. `{ adj: 25000 }`. Win over `cornerRadius`. */
  shapeAdjustments?: ShapeAdjustments;
  /** Border along the outline of the shape: hex `color`, `width` in pt. */
  border?: ImageBorder;
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
