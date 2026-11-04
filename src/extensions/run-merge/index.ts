import escape from 'regexp.escape';
import ModifyTextHelper from '../../helper/modify-text-helper';
import { XmlElement } from '../../types/xml-types';

/**
 * PowerPoint splits text into several `a:r` runs whenever it likes (spell
 * check `err="1"`, edit history, formatting changes). A tag such as `{{name}}`
 * can then be spread over runs and a per-run search never finds it.
 *
 * This moves, per paragraph, every tag that spans directly adjacent runs
 * completely into the run it starts in. Only the tag characters move: text
 * before and after the tag stays in its own run with its own formatting, and
 * runs left empty are removed. The tag (and so its replacement) takes the
 * formatting of the run it starts in.
 */
export function mergeRunsSplitByTags(
  element: XmlElement,
  openingTag: string,
  closingTag: string,
): void {
  const open = escape(openingTag);
  const close = escape(closingTag);
  const pattern = new RegExp(`${open}[^${open}${close}]+${close}`, 'g');

  const paragraphs = element.getElementsByTagName('a:p');
  for (let p = 0; p < paragraphs.length; p++) {
    adjacentRunGroups(paragraphs[p]).forEach((runs) =>
      mergeGroup(runs, pattern),
    );
  }
}

/** Runs that follow each other without another element in between. */
function adjacentRunGroups(paragraph: XmlElement): XmlElement[][] {
  const groups: XmlElement[][] = [];
  let current: XmlElement[] = [];
  const flush = () => {
    if (current.length > 1) groups.push(current);
    current = [];
  };

  for (let n = paragraph.firstChild; n; n = n.nextSibling) {
    if (n.nodeType !== 1) continue;
    const el = n as XmlElement;
    if (el.tagName === 'a:r') current.push(el);
    else if (el.tagName !== 'a:pPr' && el.tagName !== 'a:endParaRPr') flush();
  }
  flush();
  return groups;
}

function runText(run: XmlElement): string {
  return run.getElementsByTagName('a:t')[0]?.textContent ?? '';
}

function mergeGroup(runs: XmlElement[], pattern: RegExp): void {
  const texts = runs.map(runText);
  const full = texts.join('');

  // owner[pos] = index of the run that character pos ends up in.
  const owner: number[] = [];
  texts.forEach((t, r) => {
    for (let i = 0; i < t.length; i++) owner.push(r);
  });

  // A tag spanning several runs moves completely into the run it starts in.
  // Text before and after the tag stays where it was, so the boundary runs
  // keep their formatting. owner stays non-decreasing, so each run's new
  // text is still one contiguous slice of `full`.
  const changed = new Set<number>();
  for (const match of full.matchAll(pattern)) {
    const start = match.index;
    const end = start + match[0].length;
    const first = owner[start];
    if (owner[end - 1] === first) continue;
    for (let pos = start; pos < end; pos++) {
      changed.add(owner[pos]);
      owner[pos] = first;
    }
  }

  changed.forEach((r) => {
    const from = owner.indexOf(r);
    const text = from === -1 ? '' : full.slice(from, owner.lastIndexOf(r) + 1);
    const textElement = runs[r].getElementsByTagName('a:t')[0];
    if (!text) runs[r].parentNode.removeChild(runs[r]);
    else if (textElement) ModifyTextHelper.content(text)(textElement);
  });
}
