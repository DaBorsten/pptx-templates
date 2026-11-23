import { log } from '../helper/logger';
import { XmlHelper } from '../helper/xml-helper';
import { elementsByTagName } from '../helper/xml-find';
import { SlidePlaceholder } from '../types/types';
import { XmlDocument, XmlElement } from '../types/xml-types';
import type HasShapes from './has-shapes';

/**
 * Placeholder cleanup on written slides: removes duplicate placeholders,
 * strips placeholder tags without a counterpart on the target layout and
 * drops tags pptx-templates cannot process.
 *
 * Extracted from HasShapes (ROADMAP Phase 2).
 */
export class PlaceholderNormalizer {
  constructor(private host: HasShapes) {}

  /**
   * Removes all unsupported tags from slide xml and (with
   * `params.cleanupPlaceholders`) normalizes placeholder shapes.
   * E.g. added relations & tags by Thinkcell cannot be processed
   * by pptx-templates at the moment.
   */
  async cleanSlide(
    targetPath: string,
    sourcePlaceholderTypes?: SlidePlaceholder[],
  ): Promise<void> {
    const xml = await XmlHelper.getXmlFromArchive(
      this.host.targetArchive,
      targetPath,
    );

    if (this.host.cleanupPlaceholders && sourcePlaceholderTypes) {
      this.removeDuplicatePlaceholders(xml, sourcePlaceholderTypes);
      this.normalizePlaceholderShapes(xml, sourcePlaceholderTypes);
    }

    this.removeUnsupportedTags(xml);

    XmlHelper.writeXmlToArchive(this.host.targetArchive, targetPath, xml);
  }

  /**
   * Collects the placeholders present on the target slide.
   */
  async parsePlaceholders(): Promise<SlidePlaceholder[]> {
    const xml = await XmlHelper.getXmlFromArchive(
      this.host.targetArchive,
      this.host.targetPath,
    );
    const placeholderTypes: SlidePlaceholder[] = [];
    const placeholders = elementsByTagName(xml, 'p:ph');
    placeholders.forEach((placeholder: XmlElement) => {
      placeholderTypes.push({
        type: placeholder.getAttribute('type'),
        id: placeholder.getAttribute('id'),
        xml: placeholder,
      });
    });
    return placeholderTypes;
  }

  /**
   * If you insert a placeholder shape on a target slide with an empty
   * placeholder of the same type, we need to remove the existing
   * placeholder.
   */
  removeDuplicatePlaceholders(
    xml: XmlDocument,
    sourcePlaceholderTypes: SlidePlaceholder[],
  ): void {
    const placeholders = elementsByTagName(xml, 'p:ph');
    const usedTypes: Record<string, number> = {};
    placeholders.forEach((placeholder: XmlElement) => {
      const type = placeholder.getAttribute('type');
      usedTypes[type] = usedTypes[type] || 0;
      usedTypes[type]++;
    });

    for (const usedType in usedTypes) {
      const count = usedTypes[usedType];
      if (count > 1) {
        // TODO: in case more than two placeholders are of a kind,
        // this will likely remove more than intended. Should also match by id.
        const removePlaceholders = sourcePlaceholderTypes.filter(
          (sourcePlaceholder) => sourcePlaceholder.type === usedType,
        );
        removePlaceholders.forEach((removePlaceholder) => {
          const parentShapeTag = 'p:sp';
          const parentShape = XmlHelper.getClosestParent(
            parentShapeTag,
            removePlaceholder.xml,
          );
          if (parentShape) {
            XmlHelper.remove(parentShape);
          }
        });
      }
    }
  }

  /**
   * If a placeholder shape was inserted on a slide without a corresponding
   * placeholder, powerPoint will usually smash the shape's formatting.
   * This function removes the placeholder tag.
   */
  normalizePlaceholderShapes(
    xml: XmlDocument,
    sourcePlaceholderTypes: SlidePlaceholder[],
  ): void {
    const placeholders = elementsByTagName(xml, 'p:ph');
    placeholders.forEach((placeholder: XmlElement) => {
      const usedType = placeholder.getAttribute('type');
      const existingPlaceholder = sourcePlaceholderTypes.find(
        (sourcePlaceholder) => sourcePlaceholder.type === usedType,
      );
      if (!existingPlaceholder) {
        XmlHelper.remove(placeholder);
      }
    });
  }

  /**
   * Removes unsupported tags and (optionally) their now-empty parents.
   * Parents will be removed only if they become empty AND are NOT `p:nvPr`.
   */
  private removeUnsupportedTags(xml: XmlDocument): void {
    // One walk for all tags instead of a getElementsByTagName per tag
    // (each walks the whole slide). Document order, as there.
    const found = new Map<string, XmlElement[]>(
      this.host.unsupportedTags.map((tag): [string, XmlElement[]] => [tag, []]),
    );
    const stack: XmlElement[] = [];
    for (let n = xml.lastChild; n; n = n.previousSibling) {
      if (n.nodeType === 1) stack.push(n as XmlElement);
    }
    while (stack.length) {
      const el = stack.pop();
      found.get(el.tagName)?.push(el);
      for (let n = el.lastChild; n; n = n.previousSibling) {
        if (n.nodeType === 1) stack.push(n as XmlElement);
      }
    }
    // still in the document: a tag removed before can have taken it along
    const attached = (el: XmlElement) => {
      let n: unknown = el;
      while (n && n !== xml) n = (n as XmlElement).parentNode;
      return !!n;
    };

    this.host.unsupportedTags.forEach((tag, i) => {
      const drop = i === 0 ? found.get(tag) : found.get(tag).filter(attached);
      if (drop.length > 0) {
        log.debug('Cleaning unsupported tag ' + tag);

        // First get parent elements before removing
        const parents = new Set<XmlElement>();
        for (const item of drop) {
          const parent = item.parentNode as XmlElement | null;
          if (parent) {
            parents.add(parent);
          }
        }

        // Remove the unsupported tags
        drop.forEach((item) => XmlHelper.remove(item));

        // Check each parent and remove it if it has no children left
        // but never remove <p:nvPr/>
        parents.forEach((parent) => {
          if (parent.childNodes.length === 0 && parent.nodeName !== 'p:nvPr') {
            XmlHelper.remove(parent);
          }
        });
      }
    });
  }
}
