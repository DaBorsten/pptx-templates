import { Node } from '@xmldom/xmldom';
import { ArchivedFile, ArchiveType } from '../../interfaces/iarchive';
import { XmlDocument } from '../../types/xml-types';
import {
  ArchiveParams,
  AutomizerFile,
  AutomizerParams,
} from '../../types/types';
import JSZip from 'jszip';
import { patchXmldomSaxRegExpCache } from '../xmldom-sax-patch';
import { serializeXml } from '../xml-serializer';
import { XmlParseCache } from '../xml-parse-cache';

patchXmldomSaxRegExpCache();

export default class Archive {
  filename: AutomizerFile;
  params: ArchiveParams;
  buffer: Map<string, ArchivedFile> = new Map();
  options: JSZip.JSZipGeneratorOptions<'nodebuffer'> = {
    type: 'nodebuffer',
  };

  constructor(filename: AutomizerFile, params: ArchiveParams) {
    this.filename = filename;
    this.params = params;
  }

  /**
   * Parts parsed before are copied instead of parsed again. Archives that
   * read the same parts (renderDeck's) may share one cache.
   */
  parseCache = new XmlParseCache();

  parseXml(xmlString: string): XmlDocument {
    return this.parseCache.parse(xmlString);
  }

  serializeXml(xml: XmlDocument | Node) {
    return serializeXml(<Node>xml);
  }

  async writeBuffer(archiveType: ArchiveType) {
    for (const buffered of this.buffer.values()) {
      const serialized = this.serializeXml(buffered.content);
      await archiveType.write(buffered.relativePath, serialized);
    }
  }

  toBuffer(relativePath: string, content: XmlDocument): void {
    this.buffer.set(relativePath, {
      relativePath,
      name: relativePath,
      content: content,
    });
  }

  /**
   * Serializes a buffered part back into the underlying archive and drops
   * its DOM from the buffer. A parsed xmldom document costs ~25x its XML
   * source size, so parts that are finished (an appended slide after
   * cleanSlide, a master/layout after append) must not stay buffered for
   * the rest of the run. Re-reading a flushed part re-parses it from the
   * serialized content written here. No-op if the part is not buffered.
   */
  protected async flushBuffered(
    archiveType: ArchiveType,
    relativePath: string,
  ): Promise<void> {
    const buffered = this.buffer.get(relativePath);
    if (!buffered) {
      return;
    }
    await archiveType.write(relativePath, this.serializeXml(buffered.content));
    this.buffer.delete(relativePath);
  }

  setOptions(params: AutomizerParams): void {
    if (params.compression > 0) {
      this.options.compression = 'DEFLATE';
      this.options.compressionOptions = {
        level: params.compression,
      };
    }
  }

  fromBuffer(relativePath: string): ArchivedFile | undefined {
    return this.buffer.get(relativePath);
  }
}
