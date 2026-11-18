import IArchive from '../interfaces/iarchive';

import { ICounter } from '../interfaces/icounter';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { XmlHelper } from './xml-helper';

export class CountHelper implements ICounter {
  template: RootPresTemplate;
  name: string;
  count: number;

  constructor(name: string, template: RootPresTemplate) {
    this.name = name;
    this.template = template;
  }

  static increment(name: string, counters: ICounter[]): number | null {
    return CountHelper.getCounterByName(name, counters)._increment();
  }

  /**
   * Lower a counter, e.g. after slides have been removed from the archive.
   */
  static decrement(name: string, counters: ICounter[], by = 1): number {
    const counter = CountHelper.getCounterByName(name, counters);
    counter.count = Math.max(0, counter.count - by);
    return counter.count;
  }

  static count(name: string, counters: ICounter[]): number {
    return CountHelper.getCounterByName(name, counters).get();
  }

  static reset(counters: ICounter[]): void {
    counters.forEach((counter) => (counter.count = undefined));
  }

  static getCounterByName(name: string, counters: ICounter[]): ICounter {
    const counter = counters.find((c) => c.name === name);
    if (counter === undefined) {
      throw new Error(`Counter ${name} not found.`);
    }
    return counter;
  }

  _increment(): number {
    this.count++;
    return this.count;
  }

  async set(): Promise<void> {
    this.count = await this.calculateCount(this.template.archive);
  }

  get(): number {
    return this.count;
  }

  private calculateCount(presentation: IArchive): Promise<number> {
    switch (this.name) {
      case 'slides':
        return CountHelper.countSlides(presentation);
      case 'masters':
        return CountHelper.countMasters(presentation);
      case 'layouts':
        return CountHelper.countLayouts(presentation);
      case 'themes':
        return CountHelper.countThemes(presentation);
      case 'charts':
        return CountHelper.countCharts(presentation);
      case 'images':
        return CountHelper.countImages(presentation);
      case 'diagrams':
        return CountHelper.countDiagrams(presentation);
      case 'oleObjects':
        return CountHelper.countOleObjects(presentation);
    }

    throw new Error(`No way to count ${this.name}.`);
  }

  /**
   * New slides are numbered after this count. The highest slide<n>.xml
   * counts as well: a deck written by pptx-templates (removeExistingSlides +
   * cleanup) starts at slide4.xml or higher, and a new slide numbered after
   * the slide list would overwrite one of its slides.
   */
  private static async countSlides(presentation: IArchive): Promise<number> {
    const presentationXml = await XmlHelper.getXmlFromArchive(
      presentation,
      'ppt/presentation.xml',
    );
    const listed = presentationXml.getElementsByTagName('p:sldId').length;
    const files = await presentation.folder('ppt/slides');
    const highest = files.reduce((max, file) => {
      const match = /^slide(\d+)\.xml$/.exec(file.relativePath);
      return match ? Math.max(max, Number(match[1])) : max;
    }, 0);
    return Math.max(listed, highest);
  }

  private static async countMasters(presentation: IArchive): Promise<number> {
    const presentationXml = await XmlHelper.getXmlFromArchive(
      presentation,
      'ppt/presentation.xml',
    );
    return presentationXml.getElementsByTagName('p:sldMasterId').length;
  }

  /**
   * Counts the Override entries of [Content_Types].xml with the given
   * ContentType.
   */
  private static async countContentTypeOverrides(
    presentation: IArchive,
    contentType: string,
  ): Promise<number> {
    const contentTypesXml = await XmlHelper.getXmlFromArchive(
      presentation,
      '[Content_Types].xml',
    );
    const overrides = contentTypesXml.getElementsByTagName('Override');

    let count = 0;
    for (let i = 0; i < overrides.length; i++) {
      if (overrides.item(i)?.getAttribute('ContentType') === contentType) {
        count++;
      }
    }
    return count;
  }

  private static async countLayouts(presentation: IArchive): Promise<number> {
    return CountHelper.countContentTypeOverrides(
      presentation,
      'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
    );
  }

  private static async countThemes(presentation: IArchive): Promise<number> {
    return CountHelper.countContentTypeOverrides(
      presentation,
      'application/vnd.openxmlformats-officedocument.theme+xml',
    );
  }

  /**
   * The chart counter numbers chart<n>.xml, chartEx<n>.xml and their
   * style<n>/colors<n> parts alike. chartEx parts (waterfall, sunburst, …)
   * count as well, and so does the highest number taken in ppt/charts:
   * otherwise a copied chart overwrites a chart of the deck.
   */
  private static async countCharts(presentation: IArchive): Promise<number> {
    const overrides =
      (await CountHelper.countContentTypeOverrides(
        presentation,
        'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
      )) +
      (await CountHelper.countContentTypeOverrides(
        presentation,
        'application/vnd.ms-office.chartex+xml',
      ));
    const files = await presentation.folder('ppt/charts');
    const highest = files.reduce((max, file) => {
      const match = /^(?:chart|chartEx|style|colors)(\d+)\.xml$/.exec(
        file.relativePath,
      );
      return match ? Math.max(max, Number(match[1])) : max;
    }, 0);
    return Math.max(overrides, highest);
  }

  private static async countOleObjects(
    presentation: IArchive,
  ): Promise<number> {
    return CountHelper.countContentTypeOverrides(
      presentation,
      'application/vnd.openxmlformats-officedocument.oleObject',
    );
  }

  private static async countImages(presentation: IArchive): Promise<number> {
    const mediaFiles = await presentation.folder('ppt/media');
    const count = mediaFiles.filter(
      (file) => file.relativePath.indexOf('image') === 0,
    ).length;
    return count;
  }

  private static async countDiagrams(presentation: IArchive): Promise<number> {
    return CountHelper.countContentTypeOverrides(
      presentation,
      'application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml',
    );
  }
}
