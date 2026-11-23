import { createHash } from 'crypto';
import IArchive, { ArchivedFile } from '../interfaces/iarchive';
import { PptPaths } from './ppt-paths';

/**
 * PowerPoint stores each distinct media file once and lets any number of
 * relations point to the same part. pptx-templates used to copy an imported
 * image for every shape referencing it, so the same template image on ten
 * slides ended up ten times in ppt/media, bloating the output (see #145).
 *
 * This is a checksum index of the media files an output archive contains.
 * An import can look up its file contents here and re-use an identical file
 * instead of adding another copy. Relations are per shape and unaffected by
 * sharing a target, and media files are never modified in place - a media
 * file is either copied from a template or added by `Automizer.addMedia()`.
 *
 * One instance per root template; media files of source templates are
 * irrelevant, only what ends up in the output is indexed.
 */
export class MediaDeduplicator {
  archive: IArchive;

  /**
   * Checksum of a media file mapped to its filename inside ppt/media.
   */
  private files = new Map<string, string>();

  /**
   * Files added after the existing files were indexed. Looked up after
   * `files`, so an existing file wins as if it had been hashed first.
   */
  private late = new Map<string, string>();

  /**
   * Existing files not hashed yet, by size: only a file of the same size
   * can be identical, so `find` hashes just those. Reading and hashing every
   * picture of a large template took a noticeable share of each run.
   */
  private unhashed = new Map<number, ArchivedFile[]>();

  private indexed: Promise<void>;

  constructor(archive: IArchive) {
    this.archive = archive;
  }

  /**
   * Find a media file with identical contents in the output archive.
   * @param content Contents of the media file about to be copied
   * @returns Filename inside ppt/media, or undefined if it is a new file
   */
  async find(content: Buffer): Promise<string | undefined> {
    await this.indexExistingFiles();
    const sameSize = this.unhashed.get(content.length);
    if (sameSize) {
      this.unhashed.delete(content.length);
      for (const mediaFile of sameSize) await this.hash(mediaFile, this.files);
    }
    const checksum = MediaDeduplicator.checksum(content);
    return this.files.get(checksum) ?? this.late.get(checksum);
  }

  /**
   * Add a media file that has been copied into the output archive.
   */
  add(content: Buffer, filename: string): void {
    const checksum = MediaDeduplicator.checksum(content);
    const into = this.indexed ? this.late : this.files;
    if (!into.has(checksum)) {
      into.set(checksum, filename);
    }
  }

  private async hash(
    mediaFile: ArchivedFile,
    into: Map<string, string>,
  ): Promise<void> {
    // hashed late: a part removed in between is no candidate any more
    if (!this.archive.fileExists(mediaFile.name)) return;
    const content = (await this.archive.read(
      mediaFile.name,
      'nodebuffer',
    )) as Buffer;
    const checksum = MediaDeduplicator.checksum(content);
    if (!into.has(checksum)) {
      into.set(checksum, mediaFile.relativePath);
    }
  }

  /**
   * Size of a part of a jszip archive without reading it, or undefined
   * (other archive types, or not known).
   */
  private sizeOf(file: string): number | undefined {
    const zip = (this.archive as { archive?: unknown }).archive as
      { files?: Record<string, { _data?: unknown }> } | undefined;
    const data = zip?.files?.[file]?._data as
      { uncompressedSize?: unknown } | Uint8Array | undefined;
    if (data instanceof Uint8Array) return data.byteLength;
    const size = (data as { uncompressedSize?: unknown })?.uncompressedSize;
    return typeof size === 'number' && size >= 0 ? size : undefined;
  }

  /**
   * Index the media files the output archive already contains, e.g. the ones
   * coming with the root template. Runs on first use only.
   */
  private async indexExistingFiles(): Promise<void> {
    if (!this.indexed) {
      this.indexed = this.readExistingFiles();
    }
    return this.indexed;
  }

  private async readExistingFiles(): Promise<void> {
    const mediaFiles = (await this.archive.folder(PptPaths.mediaDir)).filter(
      // Directory entries have no filename to relate to
      (mediaFile) =>
        mediaFile.relativePath && !mediaFile.relativePath.endsWith('/'),
    );

    const sizes = mediaFiles.map((mediaFile) => this.sizeOf(mediaFile.name));
    if (sizes.some((size) => size === undefined)) {
      for (const mediaFile of mediaFiles)
        await this.hash(mediaFile, this.files);
      return;
    }
    mediaFiles.forEach((mediaFile, i) => {
      const bucket = this.unhashed.get(sizes[i]) ?? [];
      bucket.push(mediaFile);
      this.unhashed.set(sizes[i], bucket);
    });
  }

  static checksum(content: Buffer): string {
    return createHash('sha1').update(content).digest('hex');
  }
}
