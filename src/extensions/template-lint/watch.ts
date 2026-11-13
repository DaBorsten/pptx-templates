import fs from 'fs';
import path from 'path';

export type FileWatcher = { close: () => void };

/** Windows and macOS file systems ignore case; compare names the same way. */
const nameKey = (name: string): string =>
  process.platform === 'win32' || process.platform === 'darwin'
    ? name.toLowerCase()
    : name;

/**
 * Calls `onChange` whenever one of `files` changes. Watches the directories,
 * not the files: PowerPoint saves into a temporary file and renames it over
 * the original, and a watcher on the file itself would stay with the old,
 * deleted one. Events for other files in the directory (PowerPoint's
 * `~$deck.pptx` owner file, its `.tmp` files) are ignored.
 *
 * One save fires several events; callers debounce.
 */
export const watchFiles = (
  files: string[],
  onChange: () => void,
  onError: (error: Error) => void,
): FileWatcher => {
  const namesByDir = new Map<string, Set<string>>();
  for (const file of files) {
    const absolute = path.resolve(file);
    const dir = path.dirname(absolute);
    if (!namesByDir.has(dir)) namesByDir.set(dir, new Set());
    namesByDir.get(dir).add(nameKey(path.basename(absolute)));
  }

  const watchers: fs.FSWatcher[] = [];
  const close = () => watchers.forEach((watcher) => watcher.close());
  try {
    for (const [dir, names] of namesByDir) {
      const watcher = fs.watch(dir, (_event, filename) => {
        // Some platforms report no name; then any change may be ours
        if (!filename || names.has(nameKey(filename.toString()))) onChange();
      });
      watcher.on('error', onError);
      watchers.push(watcher);
    }
  } catch (error) {
    close();
    throw error;
  }
  return { close };
};
