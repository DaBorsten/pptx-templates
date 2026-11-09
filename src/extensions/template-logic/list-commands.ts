import fs from 'fs';
import JSZip from 'jszip';
import { DOMParser } from '@xmldom/xmldom';
import {
  isTableMarker,
  parseCommand,
  resolveAlias,
  tagSyntax,
} from './commands';
import { ListTemplateCommandsOptions, TemplateCommand } from './types';

type Element = {
  tagName: string;
  textContent: string;
  getAttribute(name: string): string;
  getElementsByTagName(name: string): ArrayLike<Element>;
  parentNode?: Element;
};

const shapeName = (paragraph: Element): string | undefined => {
  for (let node = paragraph.parentNode; node; node = node.parentNode) {
    if (node.tagName === 'p:sp' || node.tagName === 'p:graphicFrame') {
      return node.getElementsByTagName('p:cNvPr')[0]?.getAttribute('name');
    }
  }
  return undefined;
};

/**
 * Lists the template commands (`{{…}}` tags) of a .pptx, slide by slide, in
 * the spirit of docx-templates' `listCommands`. Useful to validate a
 * template or to find out which data it expects before rendering.
 *
 * @param pptx Path to a .pptx file or its contents.
 */
export const listTemplateCommands = async (
  pptx: string | Buffer | Uint8Array | ArrayBuffer,
  options?: ListTemplateCommandsOptions,
): Promise<TemplateCommand[]> => {
  const content = typeof pptx === 'string' ? fs.readFileSync(pptx) : pptx;
  const zip = await JSZip.loadAsync(content);
  const { pattern } = tagSyntax(options);

  const slides = Object.keys(zip.files)
    .map((name) => ({
      name,
      match: /^ppt\/slides\/slide(\d+)\.xml$/.exec(name),
    }))
    .filter((s) => s.match)
    .map((s) => ({ name: s.name, slide: Number(s.match[1]) }))
    .sort((a, b) => a.slide - b.slide);

  const commands: TemplateCommand[] = [];
  for (const { name, slide } of slides) {
    const xml = await zip.file(name).async('string');
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    const aliases = new Map<string, string>();
    const paragraphs = Array.from(
      doc.getElementsByTagName('a:p') as unknown as ArrayLike<Element>,
    );

    const texts = paragraphs.map((paragraph) =>
      Array.from(paragraph.getElementsByTagName('a:t'))
        .map((t) => t.textContent)
        .join(''),
    );

    // As in renderTemplate, an alias may be used before it is defined
    texts.forEach((text) => {
      for (const match of text.matchAll(pattern)) {
        try {
          const command = parseCommand(match[1], match[0]);
          if (command.type === 'ALIAS') aliases.set(command.name, command.code);
        } catch {
          // thrown below, where the tag is listed
        }
      }
    });

    paragraphs.forEach((paragraph, index) => {
      const text = texts[index];
      for (const match of text.matchAll(pattern)) {
        if (isTableMarker(match[1])) continue;
        const command = parseCommand(
          resolveAlias(match[1], aliases, match[0]),
          match[0],
        );
        commands.push({
          slide,
          shape: shapeName(paragraph),
          raw: match[0],
          type: command.type,
          code: command.code,
        });
      }
    });
  }
  return commands;
};
