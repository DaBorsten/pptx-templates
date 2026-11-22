import { XmlDocument, XmlElement } from '../../types/xml-types';
import { XmlHelper } from '../../helper/xml-helper';
import { FontRegistry } from './fonts';
import { LayoutParagraph, LayoutRun, Spacing } from './layout';

const EMU_PER_PT = 12700;

/** The parts a slide inherits its text formatting from. */
export type SlideStyles = {
  layout?: XmlDocument;
  master?: XmlDocument;
  /** Theme fonts: `+mj-lt` and `+mn-lt`. */
  majorFont?: string;
  minorFont?: string;
  /** <p:defaultTextStyle> of ppt/presentation.xml. */
  defaultTextStyle?: XmlElement;
};

/** Position and inner size of a text box, in points. */
export type BoxGeometry = {
  x: number;
  y: number;
  /** Width and height inside the insets. */
  width: number;
  height: number;
};

const children = (
  parent: XmlElement | undefined,
  tagName?: string,
): XmlElement[] => {
  if (!parent) return [];
  const result: XmlElement[] = [];
  for (let n = parent.firstChild; n; n = n.nextSibling) {
    if (
      n.nodeType === 1 &&
      (!tagName || (n as XmlElement).tagName === tagName)
    ) {
      result.push(n as XmlElement);
    }
  }
  return result;
};
const child = (parent: XmlElement | undefined, tagName: string) =>
  children(parent, tagName)[0];
const path = (parent: XmlElement | undefined, ...tags: string[]) =>
  tags.reduce<XmlElement | undefined>((el, t) => child(el, t), parent);

const numberAttr = (
  el: XmlElement | undefined,
  name: string,
): number | undefined => {
  if (!el?.hasAttribute(name)) return undefined;
  const value = Number(el.getAttribute(name));
  return Number.isFinite(value) ? value : undefined;
};
const boolAttr = (
  el: XmlElement | undefined,
  name: string,
): boolean | undefined => {
  if (!el?.hasAttribute(name)) return undefined;
  const value = el.getAttribute(name);
  return value === '1' || value === 'true';
};

type Placeholder = { type: string; idx?: string };

const placeholderOf = (
  shape: XmlElement | undefined,
): Placeholder | undefined => {
  const ph =
    path(shape, 'p:nvSpPr', 'p:nvPr', 'p:ph') ??
    path(shape, 'p:nvPicPr', 'p:nvPr', 'p:ph');
  if (!ph) return undefined;
  return {
    type: ph.getAttribute('type') || 'obj',
    idx: ph.hasAttribute('idx') ? ph.getAttribute('idx') : undefined,
  };
};

/** The shape on a layout/master that a placeholder inherits from. */
const findPlaceholder = (
  part: XmlDocument | undefined,
  ph: Placeholder,
  byType: boolean,
): XmlElement | undefined => {
  if (!part) return undefined;
  const shapes = XmlHelper.collectionToArray(part.getElementsByTagName('p:sp'));
  const typeOf = (p: Placeholder) => (p.type === 'obj' ? 'body' : p.type);
  if (!byType && ph.idx !== undefined) {
    const match = shapes.find((s) => placeholderOf(s)?.idx === ph.idx);
    if (match) return match;
  }
  return shapes.find((s) => {
    const other = placeholderOf(s);
    return other && typeOf(other) === typeOf(ph);
  });
};

/**
 * The shapes a placeholder inherits from: its counterpart on the layout,
 * then on the master. Empty for a shape that is no placeholder.
 */
export const inheritedPlaceholders = (
  shape: XmlElement,
  slide: Pick<SlideStyles, 'layout' | 'master'>,
): XmlElement[] => {
  const ph = placeholderOf(shape);
  if (!ph) return [];
  const layoutShape = findPlaceholder(slide.layout, ph, false);
  const masterShape = findPlaceholder(
    slide.master,
    placeholderOf(layoutShape) ?? ph,
    true,
  );
  return [layoutShape, masterShape].filter(Boolean);
};

const txStyleName = (ph: Placeholder): string => {
  if (ph.type === 'title' || ph.type === 'ctrTitle') return 'p:titleStyle';
  if (['body', 'obj', 'subTitle'].includes(ph.type)) return 'p:bodyStyle';
  return 'p:otherStyle';
};

/** Level style elements of a list style container (lstStyle, txStyles, …). */
const levelStyles = (
  list: XmlElement | undefined,
  level: number,
): XmlElement[] =>
  [child(list, `a:lvl${level + 1}pPr`), child(list, 'a:defPPr')].filter(
    Boolean,
  );

/**
 * Resolves the formatting of the paragraphs of one shape: paragraph and run
 * properties, the shape's list style, the placeholder it inherits from on the
 * layout and master, the master's text styles and the presentation default.
 */
export class ShapeStyles {
  private readonly ph?: Placeholder;
  private readonly layoutShape?: XmlElement;
  private readonly masterShape?: XmlElement;
  private readonly fontRef?: string;

  constructor(
    private readonly shape: XmlElement,
    private readonly slide: SlideStyles,
    private readonly fonts: FontRegistry,
  ) {
    this.ph = placeholderOf(shape);
    if (this.ph) {
      this.layoutShape = findPlaceholder(slide.layout, this.ph, false);
      this.masterShape = findPlaceholder(
        slide.master,
        placeholderOf(this.layoutShape) ?? this.ph,
        true,
      );
    }
    this.fontRef =
      path(shape, 'p:style', 'a:fontRef')?.getAttribute('idx') || undefined;
  }

  /** List styles from the shape down to the presentation default. */
  private lists(): (XmlElement | undefined)[] {
    const lstStyle = (shape?: XmlElement) =>
      path(shape, 'p:txBody', 'a:lstStyle');
    if (this.ph) {
      const master = this.slide.master
        ?.documentElement as unknown as XmlElement;
      return [
        lstStyle(this.shape),
        lstStyle(this.layoutShape),
        lstStyle(this.masterShape),
        path(master, 'p:txStyles', txStyleName(this.ph)),
        this.slide.defaultTextStyle,
      ];
    }
    return [lstStyle(this.shape), this.slide.defaultTextStyle];
  }

  private bodyPrs(): XmlElement[] {
    return [this.shape, this.layoutShape, this.masterShape]
      .map((s) => path(s, 'p:txBody', 'a:bodyPr'))
      .filter(Boolean);
  }

  private xfrm(): XmlElement | undefined {
    return [this.shape, this.layoutShape, this.masterShape]
      .map((s) => path(s, 'p:spPr', 'a:xfrm'))
      .find((x) => child(x, 'a:off') && child(x, 'a:ext'));
  }

  /** Position and inner size, or undefined when the shape has none. */
  geometry(): BoxGeometry | undefined {
    const xfrm = this.xfrm();
    if (!xfrm) return undefined;
    const off = child(xfrm, 'a:off');
    const ext = child(xfrm, 'a:ext');
    const inset = (name: string, fallback: number) =>
      (this.bodyPrs()
        .map((b) => numberAttr(b, name))
        .find((v) => v !== undefined) ?? fallback) / EMU_PER_PT;
    const x = numberAttr(off, 'x') / EMU_PER_PT;
    const y = numberAttr(off, 'y') / EMU_PER_PT;
    return {
      x,
      y,
      width:
        numberAttr(ext, 'cx') / EMU_PER_PT -
        inset('lIns', 91440) -
        inset('rIns', 91440),
      height:
        numberAttr(ext, 'cy') / EMU_PER_PT -
        inset('tIns', 45720) -
        inset('bIns', 45720),
    };
  }

  /** Reasons the text of this shape cannot be measured, if any. */
  unsupported(): string | undefined {
    const bodyPr = this.bodyPrs()[0];
    const vert = bodyPr?.getAttribute('vert');
    if (vert && vert !== 'horz') return 'vertical text';
    if (bodyPr?.getAttribute('wrap') === 'none')
      return 'text that does not wrap';
    if ((numberAttr(bodyPr, 'numCol') ?? 1) > 1)
      return 'text columns inside the shape';
    return undefined;
  }

  /** "Shrink text on overflow" as stored by PowerPoint. */
  private autofit(): { fontScale: number; lineReduction: number } {
    const fit = child(this.bodyPrs()[0], 'a:normAutofit');
    return {
      fontScale: (numberAttr(fit, 'fontScale') ?? 100000) / 100000,
      lineReduction: (numberAttr(fit, 'lnSpcReduction') ?? 0) / 100000,
    };
  }

  private typeface(name: string | undefined): string {
    const theme = (major: boolean) =>
      (major ? this.slide.majorFont : this.slide.minorFont) ?? 'Calibri';
    if (!name) return theme(this.fontRef === 'major');
    if (name.startsWith('+mj')) return theme(true);
    if (name.startsWith('+mn')) return theme(false);
    return name;
  }

  /** Paragraphs ready for layout. */
  paragraphs(paragraphs: XmlElement[]): LayoutParagraph[] {
    const { fontScale, lineReduction } = this.autofit();
    const lists = this.lists();

    return paragraphs.map((p) => {
      const pPr = child(p, 'a:pPr');
      const level = numberAttr(pPr, 'lvl') ?? 0;
      const sources = [
        pPr,
        ...lists.flatMap((l) => levelStyles(l, level)),
      ].filter(Boolean);
      const attr = (name: string) =>
        sources.map((s) => numberAttr(s, name)).find((v) => v !== undefined);
      const first = (tagName: string) =>
        sources.map((s) => child(s, tagName)).find(Boolean);
      const defRPrs = sources.map((s) => child(s, 'a:defRPr')).filter(Boolean);

      const runProps = (rPr: XmlElement | undefined) => {
        const chain = [rPr, ...defRPrs].filter(Boolean);
        const value = <T>(get: (el: XmlElement) => T | undefined) =>
          chain.map(get).find((v) => v !== undefined);
        const size = (value((el) => numberAttr(el, 'sz')) ?? 1800) / 100;
        const bold = value((el) => boolAttr(el, 'b')) ?? false;
        const italic = value((el) => boolAttr(el, 'i')) ?? false;
        const latin = value(
          (el) => child(el, 'a:latin')?.getAttribute('typeface') || undefined,
        );
        return {
          size: size * fontScale,
          face: this.fonts.face(this.typeface(latin), bold, italic),
          letterSpacing: (value((el) => numberAttr(el, 'spc')) ?? 0) / 100,
          caps: value((el) => el.getAttribute('cap') || undefined) === 'all',
        };
      };

      const runs: LayoutRun[] = [];
      children(p).forEach((node) => {
        if (node.tagName === 'a:br') {
          runs.push({ lineBreak: true, ...runProps(child(node, 'a:rPr')) });
        } else if (node.tagName === 'a:r' || node.tagName === 'a:fld') {
          const props = runProps(child(node, 'a:rPr'));
          const text = children(node, 'a:t')
            .map((t) => t.textContent ?? '')
            .join('');
          runs.push({ text: props.caps ? text.toUpperCase() : text, ...props });
        }
      });
      const end = runProps(child(p, 'a:endParaRPr'));

      const spacing = (tagName: string): Spacing | undefined => {
        const el = first(tagName);
        const pts = numberAttr(child(el, 'a:spcPts'), 'val');
        if (pts !== undefined) return { pts: pts / 100 };
        const pct = numberAttr(child(el, 'a:spcPct'), 'val');
        if (pct !== undefined) return { pct: pct / 100000 };
        return undefined;
      };
      const lineSpacing = spacing('a:lnSpc') ?? { pct: 1 };
      if (lineSpacing.pct !== undefined)
        lineSpacing.pct = Math.max(lineSpacing.pct - lineReduction, 0);

      const bulletSource = sources.find((s) =>
        ['a:buNone', 'a:buChar', 'a:buAutoNum', 'a:buBlip'].some((t) =>
          child(s, t),
        ),
      );
      let bullet: LayoutParagraph['bullet'];
      if (bulletSource && !child(bulletSource, 'a:buNone')) {
        const textSize = (runs.find((r) => 'text' in r && r.text) ?? end).size;
        const pct = numberAttr(first('a:buSzPct'), 'val');
        const pts = numberAttr(first('a:buSzPts'), 'val');
        const size =
          pts !== undefined
            ? (pts / 100) * fontScale
            : (textSize * (pct ?? 100000)) / 100000;
        const char = child(bulletSource, 'a:buChar')?.getAttribute('char');
        const buFont = first('a:buFont')?.getAttribute('typeface');
        const face = this.fonts.face(this.typeface(buFont || undefined));
        const text = char ?? (child(bulletSource, 'a:buAutoNum') ? '00.' : '•');
        bullet = { text, face, size };
      }

      return {
        runs,
        endSize: end.size,
        endFace: end.face,
        marL: (attr('marL') ?? 0) / EMU_PER_PT,
        marR: (attr('marR') ?? 0) / EMU_PER_PT,
        indent: (attr('indent') ?? 0) / EMU_PER_PT,
        tabSize: (attr('defTabSz') ?? 914400) / EMU_PER_PT,
        bullet,
        lineSpacing,
        before: spacing('a:spcBef'),
        after: spacing('a:spcAft'),
      };
    });
  }
}

/** Theme font names from a theme part. */
export const themeFonts = (
  theme: XmlDocument | undefined,
): { majorFont?: string; minorFont?: string } => {
  const scheme = theme?.getElementsByTagName('a:fontScheme')[0] as
    XmlElement | undefined;
  const latin = (tagName: string) =>
    path(scheme, tagName, 'a:latin')?.getAttribute('typeface') || undefined;
  return { majorFont: latin('a:majorFont'), minorFont: latin('a:minorFont') };
};
