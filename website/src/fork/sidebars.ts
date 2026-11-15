import type { SidebarsConfig } from '@docusaurus/plugin-content-docs';
import upstream from '../../sidebars';

// Fork extension (see FORK.md): groups the upstream sidebar into labelled
// sections without editing `website/sidebars.ts`. Pages are matched by doc
// id; anything upstream adds later and no group names lands in "More", so a
// merge can never hide a page.

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Item = any; // a sidebar item as written in sidebars.ts (string id or object)

const GROUPS: { label: string; ids: string[] }[] = [
  { label: 'Get started', ids: ['index', 'getting-started', 'concepts', 'selectors'] },
  { label: 'Modify content', ids: ['text', 'tables', 'charts', 'images', 'masters-layouts'] },
  { label: 'Templates', ids: ['template-logic', 'text-flow', 'shape-geometry', 'template-lint'] },
  { label: 'Build decks', ids: ['generation', 'hyperlinks', 'slide-management', 'helpers', 'output'] },
  { label: 'Help', ids: ['limitations', 'troubleshooting', 'testing'] },
];
const REFERENCE_LABEL = 'Reference';

const docId = (item: Item): string | undefined =>
  typeof item === 'string' ? item : item?.type === 'doc' ? item.id : undefined;

// Upstream nests some guides in plain categories ("Modify content"); flatten
// those so the groups decide the structure. Linked categories (the API
// reference) stay intact.
const flatten = (items: Item[]): Item[] =>
  items.flatMap((item) =>
    item?.type === 'category' && !item.link ? flatten(item.items) : [item],
  );

const group = (label: string, items: Item[]) => ({
  type: 'category' as const,
  label,
  collapsible: false,
  className: 'sidebar-group',
  items,
});

const upstreamItems = flatten((upstream as Record<string, Item[]>).docs);
const used = new Set<Item>();
const take = (ids: string[]) =>
  ids
    .map((id) => upstreamItems.find((item) => docId(item) === id))
    .filter((item) => item !== undefined && !used.has(item) && used.add(item));

const grouped = GROUPS.map(({ label, ids }) => group(label, take(ids))).filter(
  (g) => g.items.length > 0,
);
const reference = upstreamItems.filter((item) => item?.type === 'category' && !used.has(item));
reference.forEach((item) => used.add(item));
const rest = upstreamItems.filter((item) => !used.has(item));

const sidebars: SidebarsConfig = {
  docs: [
    ...grouped,
    ...(rest.length ? [group('More', rest)] : []),
    ...(reference.length ? [group(REFERENCE_LABEL, reference)] : []),
  ],
};

export default sidebars;
