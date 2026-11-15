import { posix } from 'path';
import IArchive from '../../interfaces/iarchive';
import { PptPaths } from '../../helper/ppt-paths';
import { XmlHelper } from '../../helper/xml-helper';
import { XmlDocument, XmlElement } from '../../types/xml-types';
import { SlideStyles, themeFonts } from './styles';

const readPart = async (
  archive: IArchive,
  file: string,
): Promise<XmlDocument | undefined> =>
  archive.fileExists(file) ? archive.readXml(file) : undefined;

/** Path of the part a relationship of `partPath` points to. */
const related = async (
  archive: IArchive,
  partPath: string,
  type: string,
): Promise<string | undefined> => {
  const rels = await readPart(archive, PptPaths.relsOf(partPath));
  if (!rels) return undefined;
  const relation = XmlHelper.collectionToArray(
    rels.getElementsByTagName('Relationship'),
  ).find((r) => r.getAttribute('Type')?.endsWith(`/${type}`));
  const target = relation?.getAttribute('Target');
  if (!target) return undefined;
  return target.startsWith('/')
    ? target.slice(1)
    : posix.normalize(posix.join(posix.dirname(partPath), target));
};

/** Layout, master, theme fonts and presentation defaults of a slide. */
export const loadSlideStyles = async (
  archive: IArchive,
  slidePath: string,
): Promise<SlideStyles> => {
  const layoutPath = await related(archive, slidePath, 'slideLayout');
  const masterPath =
    layoutPath && (await related(archive, layoutPath, 'slideMaster'));
  const themePath = masterPath && (await related(archive, masterPath, 'theme'));
  const presentation = await readPart(archive, PptPaths.presentation);
  return {
    layout: layoutPath ? await readPart(archive, layoutPath) : undefined,
    master: masterPath ? await readPart(archive, masterPath) : undefined,
    ...themeFonts(themePath ? await readPart(archive, themePath) : undefined),
    defaultTextStyle: presentation?.getElementsByTagName(
      'p:defaultTextStyle',
    )[0] as XmlElement | undefined,
  };
};
