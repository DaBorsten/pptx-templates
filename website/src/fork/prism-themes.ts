import type { PrismTheme } from 'prism-react-renderer';

// Fork extension (see FORK.md): a quiet syntax theme for the docs' TypeScript
// examples. Text stays in the neutral ink; only strings and numbers take the
// orange accent, comments step back to the secondary grey.

export const monoLight: PrismTheme = {
  plain: { color: '#1A1A1A', backgroundColor: '#FFFFFF' },
  styles: [
    { types: ['comment', 'prolog', 'doctype', 'cdata'], style: { color: '#737373' } },
    { types: ['punctuation', 'operator'], style: { color: '#6B6B6B' } },
    { types: ['keyword', 'builtin', 'important', 'atrule'], style: { color: '#111111', fontWeight: '600' } },
    { types: ['string', 'char', 'template-string', 'attr-value', 'regex', 'url'], style: { color: '#B23A0A' } },
    { types: ['number', 'boolean', 'constant', 'symbol'], style: { color: '#B23A0A' } },
    { types: ['function', 'class-name', 'maybe-class-name'], style: { color: '#111111', fontWeight: '500' } },
    { types: ['tag', 'selector', 'attr-name', 'property'], style: { color: '#3D3D3D' } },
    { types: ['variable', 'parameter'], style: { color: '#1A1A1A' } },
    { types: ['deleted'], style: { color: '#B42318' } },
    { types: ['inserted'], style: { color: '#1F7A3A' } },
  ],
};

export const monoDark: PrismTheme = {
  plain: { color: '#E5E5E5', backgroundColor: '#111111' },
  styles: [
    { types: ['comment', 'prolog', 'doctype', 'cdata'], style: { color: '#8C8C8C' } },
    { types: ['punctuation', 'operator'], style: { color: '#A3A3A3' } },
    { types: ['keyword', 'builtin', 'important', 'atrule'], style: { color: '#FFFFFF', fontWeight: '600' } },
    { types: ['string', 'char', 'template-string', 'attr-value', 'regex', 'url'], style: { color: '#FDBA74' } },
    { types: ['number', 'boolean', 'constant', 'symbol'], style: { color: '#FDBA74' } },
    { types: ['function', 'class-name', 'maybe-class-name'], style: { color: '#FFFFFF', fontWeight: '500' } },
    { types: ['tag', 'selector', 'attr-name', 'property'], style: { color: '#D4D4D4' } },
    { types: ['variable', 'parameter'], style: { color: '#E5E5E5' } },
    { types: ['deleted'], style: { color: '#F97066' } },
    { types: ['inserted'], style: { color: '#6CE9A6' } },
  ],
};
