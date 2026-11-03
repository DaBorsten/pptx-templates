import { XmlDocument, XmlElement } from '../../types/xml-types';
import { XmlHelper } from '../../helper/xml-helper';
import { PptPaths } from '../../helper/ppt-paths';
import IArchive from '../../interfaces/iarchive';
import { RootPresTemplate } from '../../interfaces/root-pres-template';
import { RenderContext, TemplateRenderer } from './renderer';
import { RenderTemplateOptions, TemplateData } from './types';

/** What slide.modify() passes as second argument (a Slide/Master/Layout). */
type PartContext = {
  targetArchive?: IArchive;
  targetPath?: string;
  targetTemplate?: RootPresTemplate;
};

const isXmlElement = (value: unknown): value is XmlElement =>
  !!value && typeof (value as XmlElement).getElementsByTagName === 'function';

/**
 * Render docx-templates style logic ({{expr}}, IF/ELSE, FOR, EXEC, ALIAS,
 * HTML, LINK, IMAGE) in all text of a shape, a table or - passed to
 * slide.modify() - a whole slide. IMAGE needs slide.modify().
 *
 * `data` may also be a Promise; functions in `data` may return Promises.
 */
export const renderTemplate =
  (
    data: TemplateData | Promise<TemplateData>,
    options?: RenderTemplateOptions,
  ) =>
  async (element: XmlDocument | XmlElement, context?: unknown): Promise<void> => {
    let write: () => void = () => undefined;
    const part = isXmlElement(context) ? undefined : (context as PartContext);
    const relsPath =
      part?.targetPath && PptPaths.relsOf(part.targetPath);

    const relations = async (): Promise<XmlElement | undefined> => {
      // modifyElement(): the slide's <Relationships> element
      if (isXmlElement(context)) return context;
      // slide.modify(): the slide/master/layout the document belongs to
      if (!part?.targetArchive || !relsPath) return undefined;
      const rels = await XmlHelper.getXmlFromArchive(part.targetArchive, relsPath);
      write = () => XmlHelper.writeXmlToArchive(part.targetArchive, relsPath, rels);
      return rels.documentElement as unknown as XmlElement;
    };

    const renderContext: RenderContext = {
      relations,
      images: async () => {
        if (!part?.targetArchive || !part.targetTemplate || !relsPath) {
          return undefined;
        }
        return {
          archive: part.targetArchive,
          template: part.targetTemplate,
          relsPath,
          relations: await relations(),
        };
      },
    };

    const renderer = new TemplateRenderer(await data, options, renderContext);
    await renderer.render(element);
    write();
  };

export { listTemplateCommands } from './list-commands';
export { TemplateRenderError } from './renderer';
export { TemplateExpressionError } from './expression';
export type {
  RenderTemplateOptions,
  TemplateData,
  TemplateCommand,
  TemplateCommandType,
  TemplateImage,
  ListTemplateCommandsOptions,
} from './types';
