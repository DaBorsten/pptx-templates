import fs from 'fs';
import path from 'path';

export type FileWatcher = { close: () => void };

/**
 * A folder to watch and which names in it count: a file to lint, or - for
 * patterns from a config file - names a new deck or folder could have.
 */
export type WatchTarget = { dir: string; accepts: (name: string) => boolean };

/** Windows and macOS file systems ignore case; compare names the same way. */
export const nameKey = (name: string): string =>
  process.platform === 'win32' || process.platform === 'darwin'
    ? name.toLowerCase()
    : name;

/** PowerPoint's owner file (`~$deck.pptx`) and temporary files. */
const isOfficeScratch = (name: string) =>
  name.startsWith('~$') || /\.tmp$/i.test(name);

/** Targets for fixed files: their folders, filtered by their names. */
export const fileTargets = (files: string[]): WatchTarget[] => {
  const namesByDir = new Map<string, Set<string>>();
  for (const file of files) {
    const absolute = path.resolve(file);
    const dir = path.dirname(absolute);
    if (!namesByDir.has(dir)) namesByDir.set(dir, new Set());
    namesByDir.get(dir).add(nameKey(path.basename(absolute)));
  }
  return [...namesByDir].map(([dir, names]) => ({
    dir,
    accepts: (name) => names.has(nameKey(name)),
  }));
};

/**
 * Calls `onChange` whenever a name a target accepts changes. Watches the
 * directories, not the files: PowerPoint saves into a temporary file and
 * renames it over the original, and a watcher on the file itself would stay
 * with the old, deleted one. PowerPoint's `~$deck.pptx` owner file and
 * `.tmp` files never count.
 *
 * One save fires several events; callers debounce.
 */
export const watchTargets = (
  targets: WatchTarget[],
  onChange: () => void,
  onError: (error: Error) => void,
): FileWatcher => {
  // Several targets for one folder: one watcher, any of them may accept
  const byDir = new Map<string, WatchTarget['accepts'][]>();
  for (const { dir, accepts } of targets) {
    const key = path.resolve(dir);
    if (!byDir.has(key)) byDir.set(key, []);
    byDir.get(key).push(accepts);
  }

  const watchers: fs.FSWatcher[] = [];
  const close = () => watchers.forEach((watcher) => watcher.close());
  try {
    for (const [dir, accepts] of byDir) {
      const watcher = fs.watch(dir, (_event, filename) => {
        // Some platforms report no name; then any change may be ours
        if (!filename) return onChange();
        const name = filename.toString();
        if (isOfficeScratch(name)) return;
        if (accepts.some((accept) => accept(name))) onChange();
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

/** Calls `onChange` whenever one of `files` changes; see `watchTargets`. */
export const watchFiles = (
  files: string[],
  onChange: () => void,
  onError: (error: Error) => void,
): FileWatcher => watchTargets(fileTargets(files), onChange, onError);
