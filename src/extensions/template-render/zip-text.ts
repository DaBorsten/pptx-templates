import JSZip from 'jszip';
import { readFastAsync } from '../../helper/archive/jszip-fast-read';

/**
 * The text of a part of `zip`, or undefined if there is none: what
 * `zip.file(name)?.async('string')` returns. A part loaded from the file
 * and unchanged since is inflated with Node's zlib right away (readFast);
 * JSZip inflates in JavaScript and waits a macrotask per part, which adds
 * up for the many small parts renderDeck reads.
 */
export const readText = async (
  zip: JSZip,
  name: string,
): Promise<string | undefined> => {
  const file = zip.file(name);
  if (!file) return undefined;
  const data = (file as unknown as { _data: unknown })._data;
  if (typeof data === 'string') {
    // set as text (zip.file(name, string)); JSZip returns it as it is
    if (!(file as unknown as { _dataBinary?: boolean })._dataBinary) {
      return data;
    }
  }
  return (
    ((await readFastAsync(file, 'string')) as string | undefined) ??
    file.async('string')
  );
};
