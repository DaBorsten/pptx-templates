import {
  ListTemplateCommandsOptions,
  RenderTemplateOptions,
} from '../template-logic/types';

/** A JSON Schema (draft 7 / 2019-09 / 2020-12), as far as the linter reads it. */
export type JsonSchema = {
  type?: string | string[];
  properties?: Record<string, JsonSchema>;
  patternProperties?: Record<string, JsonSchema>;
  additionalProperties?: boolean | JsonSchema;
  items?: JsonSchema | JsonSchema[];
  prefixItems?: JsonSchema[];
  anyOf?: JsonSchema[];
  oneOf?: JsonSchema[];
  allOf?: JsonSchema[];
  $ref?: string;
  $defs?: Record<string, JsonSchema>;
  definitions?: Record<string, JsonSchema>;
  [key: string]: unknown;
};

export type LintTemplateOptions = ListTemplateCommandsOptions &
  /** Lint as renderTemplate with the same `fixSmartQuotes` would parse. */
  Pick<RenderTemplateOptions, 'fixSmartQuotes'> & {
    /**
     * Example data. Turns on the `unknown-variable` and `unknown-property`
     * checks: every name and property path a tag reads must exist in it.
     */
    data?: Record<string, unknown>;
    /**
     * JSON Schema of the data, as an alternative (or addition) to `data`.
     * An object with `properties` is treated as closed unless it sets
     * `additionalProperties`.
     */
    schema?: JsonSchema;
    /** Same as in renderTemplate: names available besides `data`. */
    additionalJsContext?: Record<string, unknown>;
  };

export type TemplateLintSeverity = 'error' | 'warning';

export type TemplateLintRule =
  /** A tag that cannot be parsed: empty tag, invalid expression. */
  | 'syntax'
  /** `{{*name}}` without `{{ALIAS name …}}`. */
  | 'unknown-alias'
  /** `{{` without `}}` (or the other way round) in one stretch of text. */
  | 'unterminated-tag'
  /** IF/FOR not closed within the shape (or table). */
  | 'unclosed-block'
  /** END-IF/END-FOR/ELSE/ELSE-IF without a matching opening tag. */
  | 'unmatched-tag'
  /** A name that is neither in the data nor defined in the template. */
  | 'unknown-variable'
  /** A property path that does not exist in the example data or schema. */
  | 'unknown-property'
  /** FOR over a value that is not a list in the example data or schema. */
  | 'not-a-list'
  /** A tag where renderTemplate cannot render it. */
  | 'unsupported-location';

export type TemplateLintIssue = {
  severity: TemplateLintSeverity;
  rule: TemplateLintRule;
  message: string;
  /** How to fix it, when there is a typical cause. */
  hint?: string;
  /** Archive path of the part, e.g. `ppt/slides/slide2.xml`. */
  part: string;
  /** 1-based slide number (file name based), for slides and their notes. */
  slide?: number;
  /** Name of the shape that holds the tag. */
  shape?: string;
  /** The tag (or text) the issue is about. */
  tag?: string;
  /**
   * The characters of `tag` the issue is about: 0-based offsets, `end`
   * exclusive. E.g. the unterminated string, or `nmae` in `{{user.nmae}}`.
   */
  range?: { start: number; end: number };
};

export type TemplateLintResult = {
  issues: TemplateLintIssue[];
  errorCount: number;
  warningCount: number;
};
