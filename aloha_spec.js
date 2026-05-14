const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  VerticalAlign, LevelFormat, PageNumber, PageOrientation
} = require('docx');
const fs = require('fs');

// ─── SHARED STYLES ───────────────────────────────────────────────────────────
const NIGHT   = '1A160E';
const PALM    = '2B4A2A';
const BAMBOO  = 'C8A96E';
const WARM    = 'E8B86A';
const SAND    = 'F4EEE2';
const CREAM   = 'FAF7F2';
const SAND2   = 'EDE4D3';
const PALM3   = '7FA86A';
const MUTED   = '8A7A68';
const RUST    = '8B5E3C';
const BAMBOBD = '8B6E3A';

const cellBorder = (color = 'DDCFBE') => ({
  top:    { style: BorderStyle.SINGLE, size: 1, color },
  bottom: { style: BorderStyle.SINGLE, size: 1, color },
  left:   { style: BorderStyle.SINGLE, size: 1, color },
  right:  { style: BorderStyle.SINGLE, size: 1, color },
});
const noBorder = () => ({
  top:    { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  left:   { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  right:  { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
});
const accentBorder = () => ({
  top:    { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  left:   { style: BorderStyle.SINGLE, size: 12, color: BAMBOO },
  right:  { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
});
const thickBottomBorder = (color) => ({
  top:    { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  bottom: { style: BorderStyle.SINGLE, size: 8, color },
  left:   { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  right:  { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
});

// ─── HELPER FACTORIES ────────────────────────────────────────────────────────
const eyebrow = (text) => new Paragraph({
  spacing: { before: 0, after: 80 },
  children: [new TextRun({ text, font: 'Jost', size: 16, color: PALM, allCaps: true, characterSpacing: 80 })]
});

const h1 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  spacing: { before: 80, after: 160 },
  children: [new TextRun({ text, font: 'Cormorant Garamond', size: 52, color: NIGHT, bold: false })]
});

const h1italic = (regular, italic) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  spacing: { before: 80, after: 160 },
  children: [
    new TextRun({ text: regular, font: 'Cormorant Garamond', size: 52, color: NIGHT, bold: false }),
    new TextRun({ text: italic, font: 'Cormorant Garamond', size: 52, color: BAMBOBD, bold: false, italics: true })
  ]
});

const h2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 120, after: 100 },
  children: [new TextRun({ text, font: 'Cormorant Garamond', size: 40, color: NIGHT, bold: false })]
});

const h2italic = (regular, italic, extraAfter = 100) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 120, after: extraAfter },
  children: [
    new TextRun({ text: regular, font: 'Cormorant Garamond', size: 40, color: NIGHT, bold: false }),
    new TextRun({ text: italic, font: 'Cormorant Garamond', size: 40, color: BAMBOBD, bold: false, italics: true })
  ]
});

const h3 = (text, color = NIGHT) => new Paragraph({
  heading: HeadingLevel.HEADING_3,
  spacing: { before: 120, after: 80 },
  children: [new TextRun({ text, font: 'Cormorant Garamond', size: 28, color, bold: false })]
});

const body = (text, opts = {}) => new Paragraph({
  spacing: { before: 0, after: 120 },
  children: [new TextRun({
    text, font: 'Jost', size: 20, color: opts.color || MUTED,
    bold: opts.bold || false, italics: opts.italics || false
  })]
});

const bodySmall = (text, color = MUTED) => new Paragraph({
  spacing: { before: 0, after: 80 },
  children: [new TextRun({ text, font: 'Jost', size: 17, color })]
});

const mono = (text, color = PALM) => new Paragraph({
  spacing: { before: 40, after: 40 },
  children: [new TextRun({ text, font: 'Courier New', size: 17, color })]
});

const spacer = (pts = 100) => new Paragraph({
  spacing: { before: 0, after: pts },
  children: [new TextRun({ text: '' })]
});

const dividerPara = (color = BAMBOO) => new Paragraph({
  spacing: { before: 120, after: 120 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color } },
  children: [new TextRun({ text: '' })]
});

const pill = (text, bg = PALM, textColor = SAND) => new Paragraph({
  spacing: { before: 60, after: 80 },
  shading: { fill: bg, type: ShadingType.CLEAR },
  children: [new TextRun({ text: `  ${text}  `, font: 'Jost', size: 15, color: textColor, allCaps: true, characterSpacing: 60 })]
});

const sectionLabel = (text) => new Paragraph({
  spacing: { before: 200, after: 60 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 3, color: BAMBOO } },
  children: [new TextRun({ text, font: 'Jost', size: 18, color: NIGHT, bold: true })]
});

const note = (label, text) => new Paragraph({
  spacing: { before: 60, after: 60 },
  children: [
    new TextRun({ text: `${label}: `, font: 'Jost', size: 18, color: PALM, bold: true }),
    new TextRun({ text, font: 'Jost', size: 18, color: MUTED })
  ]
});

const inlineCode = (label, value) => new Paragraph({
  spacing: { before: 40, after: 40 },
  children: [
    new TextRun({ text: `${label}  `, font: 'Jost', size: 17, color: MUTED }),
    new TextRun({ text: value, font: 'Courier New', size: 17, color: PALM })
  ]
});

// Two-column key-value table helper
const kvTable = (rows, widths = [3000, 6360]) => new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: widths,
  rows: rows.map(([k, v], i) => new TableRow({
    children: [
      new TableCell({
        borders: noBorder(),
        width: { size: widths[0], type: WidthType.DXA },
        shading: { fill: i % 2 === 0 ? 'F4EEE2' : 'FAF7F2', type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 80 },
        children: [new Paragraph({ children: [new TextRun({ text: k, font: 'Jost', size: 17, color: PALM, bold: true })] })]
      }),
      new TableCell({
        borders: noBorder(),
        width: { size: widths[1], type: WidthType.DXA },
        shading: { fill: i % 2 === 0 ? 'F4EEE2' : 'FAF7F2', type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 80, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: v, font: 'Jost', size: 17, color: NIGHT })] })]
      })
    ]
  }))
});

// Highlighted block (quote / spec box)
const specBox = (lines, bg = 'F0EAD8', borderColor = BAMBOO) => new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [9360],
  rows: [new TableRow({
    children: [new TableCell({
      borders: { ...noBorder(), left: { style: BorderStyle.SINGLE, size: 16, color: borderColor } },
      shading: { fill: bg, type: ShadingType.CLEAR },
      margins: { top: 120, bottom: 120, left: 200, right: 160 },
      children: lines.map(l => new Paragraph({
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: l, font: 'Jost', size: 18, color: NIGHT })]
      }))
    })]
  })]
});

// Dark spec box
const darkBox = (lines, bg = '1A160E', borderColor = BAMBOO) => new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [9360],
  rows: [new TableRow({
    children: [new TableCell({
      borders: { ...noBorder(), left: { style: BorderStyle.SINGLE, size: 16, color: borderColor } },
      shading: { fill: bg, type: ShadingType.CLEAR },
      margins: { top: 120, bottom: 120, left: 200, right: 160 },
      children: lines.map(l => new Paragraph({
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: l, font: 'Jost', size: 18, color: 'F4EEE2' })]
      }))
    })]
  })]
});

// Three-column table helper
const threeColTable = (headers, rows) => {
  const W = [2200, 3200, 3960];
  const headerRow = new TableRow({
    children: headers.map((h, i) => new TableCell({
      borders: { ...noBorder(), bottom: { style: BorderStyle.SINGLE, size: 6, color: BAMBOO } },
      width: { size: W[i], type: WidthType.DXA },
      shading: { fill: NIGHT, type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 100, left: 120, right: 80 },
      children: [new Paragraph({ children: [new TextRun({ text: h, font: 'Jost', size: 16, color: SAND, bold: true, allCaps: true, characterSpacing: 40 })] })]
    }))
  });
  const dataRows = rows.map((r, ri) => new TableRow({
    children: r.map((cell, i) => new TableCell({
      borders: { ...noBorder(), bottom: { style: BorderStyle.SINGLE, size: 2, color: SAND2 } },
      width: { size: W[i], type: WidthType.DXA },
      shading: { fill: ri % 2 === 0 ? 'FAF7F2' : 'F4EEE2', type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 100, left: 120, right: 80 },
      children: [new Paragraph({ children: [new TextRun({ text: cell, font: i === 0 ? 'Courier New' : 'Jost', size: 17, color: i === 0 ? PALM : NIGHT })] })]
    }))
  }));
  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: W, rows: [headerRow, ...dataRows] });
};

// ─── COVER PAGE ──────────────────────────────────────────────────────────────
const coverSection = [
  spacer(400),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 60 },
    children: [new TextRun({ text: 'ALOHA CAFÉ', font: 'Jost', size: 48, color: NIGHT, bold: true, allCaps: true, characterSpacing: 160 })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 40 },
    children: [new TextRun({ text: 'The Back Yard · Vijayawada', font: 'Cormorant Garamond', size: 28, color: BAMBOBD, italics: true })]
  }),
  dividerPara(BAMBOO),
  spacer(80),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 60 },
    children: [new TextRun({ text: 'PREMIUM WEBSITE SPECIFICATION', font: 'Jost', size: 22, color: MUTED, allCaps: true, characterSpacing: 120 })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 40 },
    children: [new TextRun({ text: 'Optimised for Greta.sh AI Website Builder', font: 'Jost', size: 20, color: MUTED })]
  }),
  spacer(60),
  new Table({
    width: { size: 5400, type: WidthType.DXA },
    columnWidths: [2700, 2700],
    rows: [
      new TableRow({ children: [
        new TableCell({ borders: noBorder(), width: { size: 2700, type: WidthType.DXA }, shading: { fill: 'F0EAD8', type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 140, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: 'Document Type', font: 'Jost', size: 16, color: MUTED })] })] }),
        new TableCell({ borders: noBorder(), width: { size: 2700, type: WidthType.DXA }, shading: { fill: 'F0EAD8', type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 80, right: 140 }, children: [new Paragraph({ children: [new TextRun({ text: 'AI Builder Spec v1.0', font: 'Jost', size: 16, color: NIGHT })] })] })
      ]}),
      new TableRow({ children: [
        new TableCell({ borders: noBorder(), width: { size: 2700, type: WidthType.DXA }, shading: { fill: 'FAF7F2', type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 140, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: 'Target Platform', font: 'Jost', size: 16, color: MUTED })] })] }),
        new TableCell({ borders: noBorder(), width: { size: 2700, type: WidthType.DXA }, shading: { fill: 'FAF7F2', type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 80, right: 140 }, children: [new Paragraph({ children: [new TextRun({ text: 'Greta.sh', font: 'Jost', size: 16, color: NIGHT })] })] })
      ]}),
      new TableRow({ children: [
        new TableCell({ borders: noBorder(), width: { size: 2700, type: WidthType.DXA }, shading: { fill: 'F0EAD8', type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 140, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: 'Brand Tone', font: 'Jost', size: 16, color: MUTED })] })] }),
        new TableCell({ borders: noBorder(), width: { size: 2700, type: WidthType.DXA }, shading: { fill: 'F0EAD8', type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 80, right: 140 }, children: [new Paragraph({ children: [new TextRun({ text: 'Cinematic · Tropical · Premium', font: 'Jost', size: 16, color: NIGHT })] })] })
      ]}),
      new TableRow({ children: [
        new TableCell({ borders: noBorder(), width: { size: 2700, type: WidthType.DXA }, shading: { fill: 'FAF7F2', type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 140, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: 'Location', font: 'Jost', size: 16, color: MUTED })] })] }),
        new TableCell({ borders: noBorder(), width: { size: 2700, type: WidthType.DXA }, shading: { fill: 'FAF7F2', type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 80, right: 140 }, children: [new Paragraph({ children: [new TextRun({ text: 'Vijayawada, Andhra Pradesh, India', font: 'Jost', size: 16, color: NIGHT })] })] })
      ]}),
    ]
  }),
  spacer(200),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: 'Creative Direction by Claude · Anthropic', font: 'Jost', size: 16, color: 'CCBFAE', italics: true })]
  }),
  spacer(100),
];

// ─── SECTION 1: BRAND OVERVIEW ───────────────────────────────────────────────
const brandSection = [
  sectionLabel('SECTION 1 — BRAND & DESIGN OVERVIEW'),
  spacer(60),
  h1italic('Design Philosophy: ', 'Atmosphere Over Everything'),
  body('Aloha Café is not a generic restaurant website. It is a cinematic, immersive digital experience that mirrors the feeling of stepping through a bamboo gate into a warm garden escape. Every design decision — font weight, colour, spacing, copy length, image crop — must answer one question: does this feel like a place I want to sit in for two hours?'),
  spacer(60),
  h2('Brand Personality Matrix'),
  kvTable([
    ['Primary Tone',     'Cinematic · Calm · Premium'],
    ['Secondary Tone',   'Warm · Intimate · Unhurried'],
    ['Visual Register',  'Boutique tropical lounge — not a beach bar, not a chain café'],
    ['Copy Voice',       'Minimal, poetic, confident. No exclamation marks. No "amazing" or "delicious."'],
    ['Avoid',           'Busy layouts · generic sans fonts · bright primary colours · cheesy tropical clichés'],
    ['Inspired By',      'Aesop stores · Soho House · Sézane editorial pages'],
  ]),
  spacer(80),
  h2('Target Audience'),
  kvTable([
    ['Primary',     'Young professionals (22–32) · couples on date nights'],
    ['Secondary',   'College students · content creators · remote workers'],
    ['Psychograph', 'Aesthetics-aware · Instagram-native · value experience over price'],
    ['Device Split','60% mobile · 35% desktop · 5% tablet — design mobile-first'],
  ]),
  spacer(80),
  sectionLabel('SECTION 2 — DESIGN SYSTEM'),
  spacer(60),
  h1italic('Colour Palette: ', 'Extracted From the Space'),
  body('All colours are derived directly from the physical café — the bamboo walls, warm LED strips, coconut palms, night garden, and sand-coloured architecture.'),
  spacer(60),
  threeColTable(
    ['Token', 'Hex Value', 'Usage Rule'],
    [
      ['--night',      '#1A160E',   'Hero backgrounds · footer · dark sections · primary text on light'],
      ['--palm',       '#2B4A2A',   'CTA buttons · interactive elements · nav active state'],
      ['--palm-mid',   '#3D6640',   'Button hover state · links'],
      ['--palm-light', '#7FA86A',   'Eyebrow labels · small accents only — never dominant'],
      ['--bamboo',     '#C8A96E',   'Section headlines · price text · icon strokes · gold accent'],
      ['--warm',       '#E8B86A',   'Italic serif moments · hover glow · LED warmth effect'],
      ['--sand',       '#F4EEE2',   'Body text on dark backgrounds · card fills on light pages'],
      ['--sand2',      '#EDE4D3',   'Dividers · card borders · input borders'],
      ['--cream',      '#FAF7F2',   'Primary page background · light section backgrounds'],
      ['--bamboo-d',   '#8B6E3A',   'Italic headline text · section headings with serif italic'],
      ['--muted',      '#8A7A68',   'Body copy · descriptive text · metadata'],
      ['--rust',       '#8B5E3C',   'Ghost CTA text · secondary link colour'],
    ]
  ),
  spacer(100),
  specBox([
    'CRITICAL COLOUR RULES:',
    '',
    '1. Never use pure white (#FFFFFF) as a background — always use --cream or --sand.',
    '2. Never use pure black for text — always use --night (#1A160E).',
    '3. --palm-light (#7FA86A) must never cover more than 5% of any section.',
    '4. The --warm colour (#E8B86A) is reserved for italic serif text and hover states ONLY.',
    '5. Dark sections use --night background + --sand text. Never invert this.',
    '6. All section transitions should shift between adjacent palette values — never jump from --cream to --night without a sand intermediate.',
  ], 'F0EAD8', BAMBOO),
  spacer(100),
  h2('Typography System'),
  body('Two typefaces only. No system fonts. No fallback to Arial or Inter.'),
  spacer(60),
  kvTable([
    ['Display Font',     'Cormorant Garamond — weight 300 (Light) and 300 Italic only'],
    ['UI / Body Font',   'Jost — weight 300, 400, 500'],
    ['Import URL',       'Google Fonts: Cormorant Garamond:ital,wght@0,300;1,300 + Jost:wght@300;400;500'],
    ['Never Use',        'Playfair Display · Montserrat · Inter · Roboto · system-ui as primary'],
  ]),
  spacer(80),
  threeColTable(
    ['Element', 'Specification', 'Notes'],
    [
      ['H1 — Hero Headline',     'Cormorant Garamond · 300 · 4.8–5.5rem · line-height 1.05',   'Clamp: clamp(2.8rem, 6vw, 5.5rem). Key word in italic + --bamboo-d colour'],
      ['H2 — Section Title',     'Cormorant Garamond · 300 · clamp(2rem, 4vw, 3rem)',           'Line-height 1.15. Italic variant in --bamboo-d for emotional words'],
      ['H3 — Card / Sub-title',  'Cormorant Garamond · 400 · 1.2–1.35rem',                     'No bold. Upright weight for card names, italic for pull-quotes'],
      ['Eyebrow Label',          'Jost · 400 · 0.67rem · letter-spacing 0.3em · ALL CAPS',     'Colour: --palm-light. Used above every section headline'],
      ['Body Copy',              'Jost · 300 · 0.9–0.95rem · line-height 1.85',                '--muted colour. Max-width 440px for readability'],
      ['UI Labels / Nav',        'Jost · 400 · 0.7–0.72rem · letter-spacing 0.2em · ALL CAPS', '--muted on light / rgba(sand,0.45) on dark backgrounds'],
      ['Price Text',             'Cormorant Garamond · 400 · 1.1rem',                           '--bamboo-d colour only'],
      ['CTA Button',             'Jost · 500 · 0.72–0.75rem · letter-spacing 0.16em · ALL CAPS','Never serif in buttons'],
      ['Footer Micro',           'Jost · 300 · 0.65rem · letter-spacing 0.1em',                'rgba(sand, 0.2) on --night background'],
    ]
  ),
  spacer(100),
  h2('Spacing System'),
  body('Use a generous, editorial spacing scale. Sections must breathe. Never crowd content.'),
  kvTable([
    ['Section vertical padding',    '6rem top + 6rem bottom (desktop) · 4rem top + 4rem bottom (mobile)'],
    ['Content max-width',           '1100px centred — never full-bleed text'],
    ['Column gap (2-col layouts)',  '5rem (desktop) · stacked vertically on mobile'],
    ['Card gap',                    '1.2–1.5rem'],
    ['Paragraph spacing',           'margin-bottom: 1.2rem between body paragraphs'],
    ['Section eyebrow → headline', '0.8rem gap'],
    ['Headline → body copy',       '1.2rem gap'],
    ['Body copy → CTA',            '1.8–2rem gap'],
    ['Mobile section padding',     'min 3.5rem horizontal padding — never less than 1.5rem'],
  ]),
];

// ─── SECTION 3: NAVBAR ───────────────────────────────────────────────────────
const navSection = [
  sectionLabel('SECTION 3 — NAVIGATION'),
  spacer(60),
  h1italic('Navbar: ', 'Fixed · Dark · Minimal'),
  spacer(40),
  specBox([
    'NAVBAR SPECIFICATION',
    '',
    'Position:          position: fixed; top: 0; left: 0; right: 0; z-index: 100',
    'Default State:     background: rgba(26,22,14, 0.88) + backdrop-filter: blur(16px)',
    'On Scroll:         background: rgba(26,22,14, 0.96) — slightly more opaque',
    'Height:            70px desktop · 60px mobile',
    'Border:            border-bottom: 1px solid rgba(200,169,110, 0.12)',
    'Transition:        background 0.3s ease, border-color 0.3s ease',
  ], 'F0EAD8', BAMBOO),
  spacer(80),
  h2('Logo (Left-Aligned)'),
  kvTable([
    ['Type',        'Wordmark — NO icon/symbol in the nav bar'],
    ['Text',        '"Aloha" + italic "Café"'],
    ['Font',        'Cormorant Garamond · 300 · 1.5rem · letter-spacing 0.1em'],
    ['"Aloha"',     'Colour: --sand (#F4EEE2)'],
    ['"Café"',      'Colour: --warm (#E8B86A) · italic'],
    ['Hover',       'opacity: 0.85 transition 0.2s'],
    ['Mobile',      'Same wordmark, reduced to 1.3rem'],
  ]),
  spacer(80),
  h2('Navigation Links (Centre / Right)'),
  kvTable([
    ['Items',            'Menu · Our Space · Gallery · Visit'],
    ['Font',             'Jost · 400 · 0.7rem · letter-spacing 0.2em · ALL CAPS'],
    ['Default Colour',   'rgba(244,238,226, 0.45) — muted sand on dark nav'],
    ['Hover Colour',     '--bamboo (#C8A96E) · transition 0.2s'],
    ['Active/Current',   '--bamboo · no underline · no bold — colour change only'],
    ['Gap Between Links','2.5rem desktop'],
    ['Mobile',           'Hide links · show hamburger icon (see mobile section)'],
  ]),
  spacer(80),
  h2('CTA Button (Far Right)'),
  kvTable([
    ['Text',             '"Reserve a Table"'],
    ['Background',       '--palm (#2B4A2A)'],
    ['Text Colour',      '--sand (#F4EEE2)'],
    ['Font',             'Jost · 500 · 0.68rem · letter-spacing 0.15em · ALL CAPS'],
    ['Shape',            'border-radius: 2rem (pill)'],
    ['Padding',          '0.5rem 1.4rem'],
    ['Hover',            'background: --palm-mid (#3D6640) · transform: translateY(-1px)'],
    ['Mobile',           'Hidden from nav — replaced by sticky bottom bar CTA'],
  ]),
];

// ─── SECTION 4: HOMEPAGE SECTIONS ─────────────────────────────────────────────
const homepageSection = [
  sectionLabel('SECTION 4 — HOMEPAGE STRUCTURE (IN ORDER)'),
  spacer(60),
  h1italic('10 Sections. ', 'Zero Clutter.'),
  body('Build the homepage in exactly this sequence. Do not reorder. Each section has a distinct background colour to create natural visual rhythm without needing decorative dividers.'),
  spacer(80),

  // SECTION 4.1 HERO
  h2('4.1 — HERO SECTION'),
  dividerPara(BAMBOO),
  kvTable([
    ['Background',         'Full-bleed Image 2 (bamboo corridor at night) as CSS background-image'],
    ['Overlay',            'linear-gradient(to top, rgba(26,22,14,0.95) 0%, rgba(26,22,14,0.5) 40%, rgba(26,22,14,0.15) 100%)'],
    ['Min-height',         '100vh desktop · 92vh mobile'],
    ['Content position',   'Content anchored bottom-left: align-items: flex-end; padding: 0 3rem 5rem'],
    ['Content max-width',  '600px'],
  ]),
  spacer(60),
  darkBox([
    'EYEBROW TEXT:   Aloha — The Back Yard · Vijayawada',
    'EYEBROW STYLE:  Jost · 0.67rem · letter-spacing 0.35em · --bamboo colour · ALL CAPS · margin-bottom 1.2rem',
    '',
    'HEADLINE:       "Where the city exhales."',
    'HEADLINE STYLE: Cormorant Garamond · 300 · clamp(2.8rem, 6vw, 5rem) · --sand colour',
    '                "exhales." must be on its own line, in italic, in --warm colour',
    '',
    'SUBHEADLINE:    "A garden café tucked behind bamboo walls and swaying palms.',
    '                Open-air seating, warm evenings, and a menu worth lingering over."',
    'SUB STYLE:      Jost · 300 · 0.95rem · rgba(sand, 0.55) · line-height 1.8 · max-width 400px',
    '',
    'CTA PRIMARY:    "Find Your Table" — pill button, --bamboo background, --night text',
    'CTA GHOST:      "Explore the Space →" — no bg, rgba(sand,0.6) colour, text underline only',
    'BUTTON GAP:     1.2rem between CTAs · flex-direction row',
  ], '1A160E', BAMBOO),
  spacer(80),
  h3('Hero Animations'),
  kvTable([
    ['Eyebrow',         'opacity: 0 → 1, translateY(18px) → 0 · duration 0.8s · delay 0.3s · ease'],
    ['Headline',        'opacity: 0 → 1, translateY(18px) → 0 · duration 0.9s · delay 0.5s · ease'],
    ['Subheadline',     'opacity: 0 → 1, translateY(18px) → 0 · duration 0.9s · delay 0.7s · ease'],
    ['Buttons',         'opacity: 0 → 1, translateY(18px) → 0 · duration 0.9s · delay 0.9s · ease'],
    ['Badge (top right)','opacity: 0 → 1 · duration 1s · delay 1.3s (no movement)'],
    ['Scroll indicator','opacity: 0 → 1 · duration 1s · delay 1.5s · a gentle loop pulse animation'],
  ]),
  spacer(60),
  h3('Hero Additional Elements'),
  body('Circular badge (top-right, inside hero): 100×100px circle · border: 1px solid rgba(sand, 0.12) · contains star symbol + "Vijayawada\'s Garden Café" in tiny tracked Jost. Decorative only.'),
  body('Scroll indicator (bottom-centre): "SCROLL" text in 0.58rem Jost + a 40px vertical line with gradient fade. Subtle looping pulse animation.'),
  body('SVG bamboo silhouette art on right edge: hand-drawn style bamboo stalks and fronds in rgba(bamboo, 0.15). Decorative layer behind content, above hero image.'),
  spacer(100),

  // SECTION 4.2 ETHOS
  h2('4.2 — ETHOS STRIP'),
  dividerPara(BAMBOO),
  kvTable([
    ['Background',         '--sand (#F4EEE2)'],
    ['Layout',             '3 equal columns (grid-template-columns: repeat(3,1fr))'],
    ['Column dividers',    'border-right: 1px solid --sand2 on first two columns'],
    ['Padding',            '4rem 2.5rem per column'],
    ['Purpose',            'Establish the three moods of the café — no images, typography only'],
  ]),
  spacer(60),
  specBox([
    'COLUMN 1 — "01 · The Back Yard"',
    'Number:  "01" · Cormorant · 2.5rem · --bamboo · opacity 0.5',
    'Title:   "The Back Yard" · Cormorant · 1.2rem · --night',
    'Body:    "An open-air terrace where afternoon light filters through bamboo blinds"',
    '',
    'COLUMN 2 — "02 · The Garden"',
    'Number:  "02" · same style',
    'Title:   "The Garden" · Cormorant · 1.2rem · --night',
    'Body:    "A bamboo-walled corridor alive with green plantings and warm globe pendants"',
    '',
    'COLUMN 3 — "03 · The Evening"',
    'Number:  "03" · same style',
    'Title:   "The Evening" · Cormorant · 1.2rem · --night',
    'Body:    "Golden hour at Aloha hits differently. Warm lights, cooler air, nowhere else to be."',
  ], 'F0EAD8', BAMBOO),
  spacer(100),

  // SECTION 4.3 ABOUT
  h2('4.3 — ABOUT / OUR SPACE'),
  dividerPara(BAMBOO),
  kvTable([
    ['Background',         '--cream (#FAF7F2)'],
    ['Layout',             'Two columns: 1fr 1fr · gap 5rem · align-items: center'],
    ['Left column',        'Text content'],
    ['Right column',       'Photo (Image 3: dusk exterior with coconut palm)'],
    ['Section padding',    '6rem 2.5rem'],
    ['Image treatment',    'border-radius: 1rem · overflow: hidden · height: 440px · object-fit: cover'],
    ['Image crop',         'Crop Image 3 to exclude neighbouring building. Focus on palm + lit veranda'],
  ]),
  spacer(60),
  specBox([
    'EYEBROW:     "Our Space"',
    'HEADLINE:    "Not just a café." (line break) "A whole other world."',
    '             "other world." in italic, --bamboo-d colour',
    '',
    'BODY P1:     "Tucked behind bamboo walls and a pair of coconut palms, Aloha Café is',
    '             Vijayawada\'s most unhurried address. A converted villa with an open courtyard,',
    '             a garden corridor lit after dark, and a verandah that slows time on its own."',
    '',
    'BODY P2:     "Whether you arrive with a laptop or someone worth talking to, the space',
    '             accommodates both — without trying too hard."',
    '',
    'GHOST CTA:   "Our Story →" · --palm-mid colour · text underline on bottom · no bg',
    '',
    'IMAGE BADGE (absolute, top-right): pill "Vijayawada" · bg --palm · text --sand · 0.65rem Jost',
    'IMAGE TAG (absolute, bottom-left): "Bamboo · Garden · Open-air" · bg rgba(sand,0.92)',
  ], 'F0EAD8', BAMBOO),
  spacer(100),

  // SECTION 4.4 MENU
  h2('4.4 — MENU SHOWCASE'),
  dividerPara(BAMBOO),
  kvTable([
    ['Background',         '--sand (#F4EEE2)'],
    ['Layout',             'Centred header · category tabs · 3-column card grid'],
    ['Section padding',    '6rem 2.5rem'],
    ['Card grid',          'grid-template-columns: repeat(3,1fr) · gap 1.2rem · max-width 860px · centred'],
    ['No food photos',     'Typography-only cards — see strategy in Section 9'],
  ]),
  specBox([
    'EYEBROW:     "The Menu" · centred',
    'HEADLINE:    "A menu that earns" (line break) "its moment."',
    '             "its moment." italic, --bamboo-d',
    'SUBTEXT:     "Small plates, fresh drinks, and a few things you\'ll come back for."',
    '             max-width 380px · centred',
    '',
    'CATEGORY TABS: pill buttons · All · Drinks · Bites · Specials',
    'Tab default:   transparent bg · --sand2 border · --muted text · 0.7rem Jost ALL CAPS',
    'Tab active:    --palm bg · --palm border · --sand text',
    'Tab hover:     same as active',
    '',
    'BOTTOM CTA:  "View Full Menu" · --bamboo bg · --night text · pill · centred',
  ], 'F0EAD8', BAMBOO),
  spacer(80),
  h3('Menu Card Specification'),
  threeColTable(
    ['Card Element', 'Style', 'Content Example'],
    [
      ['Card container',    'bg --cream · border 1px --sand2 · radius 1rem · padding 1.8rem',       '—'],
      ['Card hover',        'transform: translateY(-3px) · top border: 3px gradient palm→bamboo',   '—'],
      ['Item name',         'Cormorant Garamond · 1.2rem · --night · line-height 1.2',              '"Cold Brew Ritual"'],
      ['Descriptor',        'Jost · 0.77rem · --muted · line-height 1.65 · margin-bottom 1rem',     '"Slow-dripped overnight, ice-settled."'],
      ['Price',             'Cormorant Garamond · 1.1rem · --bamboo-d',                             '"₹180"'],
      ['Optional badge',    '0.58rem Jost ALL CAPS · rgba(palm, 0.1) bg · --palm-mid text',         '"Bestseller" / "New"'],
    ]
  ),
  spacer(100),

  // SECTION 4.5 EXPERIENCE
  h2('4.5 — EXPERIENCE / WHY ALOHA'),
  dividerPara(BAMBOO),
  kvTable([
    ['Background',     '--night (#1A160E) — the dark anchor section'],
    ['Layout',         'Centred header + 3-column card grid · margin-top 3.5rem'],
    ['Card border',    '1px solid rgba(sand, 0.08) · radius 1rem · padding 2.2rem 1.8rem'],
    ['Card hover',     'border-color: rgba(200,169,110, 0.3) · transition 0.25s'],
    ['Purpose',        'Communicate the three pillars of why people visit: Garden · Work · Date Night'],
  ]),
  specBox([
    'EYEBROW:     "Why Aloha" · --palm-light',
    'HEADLINE:    "Three reasons to stay" (line break) "longer than you planned."',
    '             "longer than you planned." italic · --warm colour',
    '',
    'CARD 1 — The Garden After Dark',
    '  Number: "01" · Cormorant 2.8rem · --bamboo · opacity 0.4',
    '  Title:  "The Garden After Dark" · Cormorant · 1.25rem · --sand',
    '  Body:   "A bamboo-walled corridor lit by warm LED strips and globe pendants.',
    '           The kind of space that makes your photos look effortless."',
    '  Body style: Jost · 0.8rem · rgba(sand, 0.4) · line-height 1.75',
    '',
    'CARD 2 — The Working Afternoon',
    '  Title:  "The Working Afternoon"',
    '  Body:   "Good wifi, unhurried service, and an atmosphere that does not rush you.',
    '           Come with a deadline. Leave without one."',
    '',
    'CARD 3 — The Date Night Setting',
    '  Title:  "The Date Night Setting"',
    '  Body:   "Warm light, bamboo walls, and a menu that does not distract from the',
    '           conversation. Aloha does the rest."',
  ], 'F0EAD8', BAMBOO),
  spacer(100),

  // SECTION 4.6 GALLERY
  h2('4.6 — GALLERY SECTION'),
  dividerPara(BAMBOO),
  kvTable([
    ['Background',         '--cream (#FAF7F2)'],
    ['Layout',             'Asymmetric mosaic grid — NOT a uniform grid'],
    ['Grid specification', 'grid-template-columns: 2fr 1fr 1fr · grid-template-rows: 220px 220px'],
    ['First cell',         'grid-row: 1/3 — spans full height (440px total)'],
    ['Gap',                '0.8rem'],
    ['Border-radius',      'border-radius: 1rem on grid container · overflow: hidden'],
    ['Cell hover',         'transform: scale(1.04) on inner fill · transition 0.4s ease'],
  ]),
  specBox([
    'CELL 1 (Large, spans 2 rows):   Image 2 — bamboo corridor at night',
    '  Label overlay: "The Garden Corridor" — 0.6rem Jost ALL CAPS · rgba(sand,0.35)',
    '',
    'CELL 2 (top right):             Image 3 — dusk exterior with palm',
    '  Label overlay: "Golden Hour"',
    '',
    'CELL 3 (middle right):          Image 4 — daytime arrival/entrance view',
    '  Label overlay: "The Entrance"',
    '',
    'CELL 4 (bottom right top):      Image 1 cropped — pendant lights close-up',
    '  Label overlay: "Inside the Yard"',
    '',
    'CELL 5 (bottom right bottom):   UGC / Instagram embedded or placeholder',
    '  Label overlay: "Your Moments"',
    '',
    'BELOW GRID: "Follow us on Instagram · @alohacafe.vjw"',
    '  Style: Jost · 0.72rem · --muted · centred · margin-top 1.8rem',
    '  Handle link: --palm-mid colour · border-bottom 1px --palm-light',
  ], 'F0EAD8', BAMBOO),
  spacer(100),

  // SECTION 4.7 REVIEWS
  h2('4.7 — REVIEWS SECTION'),
  dividerPara(BAMBOO),
  kvTable([
    ['Background',     '--sand (#F4EEE2)'],
    ['Layout',         '3-column card grid · margin-top 3rem'],
    ['Card style',     'bg --cream · border 1px --sand2 · radius 1rem · padding 1.8rem'],
    ['No photos',      'Use initials avatar (36×36px circle, --sand2 bg) — not real faces'],
  ]),
  specBox([
    'EYEBROW:     "What People Say" · centred',
    'HEADLINE:    "Heard around" (line break) "the garden." · italic on "the garden."',
    '',
    'REVIEW 1 — Couple',
    '  Stars:  5 gold stars (★★★★★) · --bamboo colour · 0.85rem',
    '  Quote:  "We came for coffee and stayed for three hours.',
    '           The bamboo garden at night is unlike anything else in Vijayawada."',
    '  Quote style: Cormorant Garamond · 1.02rem · 300 · italic · --night · line-height 1.7',
    '  Name:   "Rohan & Priya" · Jost 0.8rem · 500 · --night',
    '  Type:   "Couple · Regular guests" · Jost 0.68rem · --muted',
    '',
    'REVIEW 2 — Creator',
    '  Quote:  "Every corner photographs beautifully.',
    '           Finally a café that understands atmosphere over aesthetics."',
    '  Name:   "Sahithi R." · Type: "Content Creator"',
    '',
    'REVIEW 3 — Student',
    '  Quote:  "My go-to for focusing and unwinding in the same afternoon.',
    '           The vibe here just makes everything easier."',
    '  Name:   "Aakash T." · Type: "Student · SRM University"',
  ], 'F0EAD8', BAMBOO),
  spacer(100),

  // SECTION 4.8 VISIT / RESERVE
  h2('4.8 — VISIT & RESERVE'),
  dividerPara(BAMBOO),
  kvTable([
    ['Background',     '--palm (#2B4A2A) — warm dark green section'],
    ['Layout',         '2 columns: 1fr 1fr · gap 4rem · align-items: center'],
    ['Left column',    'Address, hours, best-time info'],
    ['Right column',   'Reservation form card'],
    ['Purpose',        'Primary conversion section — this is where visitors become guests'],
  ]),
  specBox([
    'LEFT COLUMN',
    'EYEBROW:     "Find Us" · --palm-light',
    'HEADLINE:    "Come as you are." (line break) "Stay as long as you like."',
    '             "Stay as long as you like." italic · --warm colour',
    '',
    'INFO ITEMS (icon + label + value):',
    '  📍 Address:    "Aloha — The Back Yard · Vijayawada, Andhra Pradesh"',
    '  🕐 Hours:      "Mon–Sun · 10:00 AM – 11:00 PM · Kitchen closes 10:30 PM"',
    '  🌿 Best Visit: "Evenings after 6 PM for the full garden lighting experience"',
    '',
    'ICON style: 32×32px circle · bg rgba(sand,0.08) · contains emoji/SVG icon',
    'Label style: Jost · 0.67rem · ALL CAPS · letter-spacing 0.15em · rgba(sand,0.4)',
    'Value style: Jost · 0.88rem · rgba(sand,0.8) · 300 weight',
    '',
    'RIGHT COLUMN — Reserve Form Card',
    'Card bg:    rgba(sand, 0.06) · border: 1px solid rgba(sand,0.1) · radius 1rem · padding 2rem',
    'Form title: "Reserve a table" · Cormorant · 1.4rem · --sand',
    'Form sub:   "Or message directly on WhatsApp — we\'ll confirm in minutes." · --muted',
    '',
    'FIELDS: Name (text) · Date (date picker) · Guests (number) — 3 fields maximum',
    'Input style: bg rgba(sand,0.07) · border rgba(sand,0.12) · --sand text · radius 0.5rem',
    'Input focus: border-color rgba(bamboo,0.4) · no glow/box-shadow',
    '',
    'CTA BUTTON: "📲 Confirm on WhatsApp" · --bamboo bg · --night text · full width · pill',
  ], 'F0EAD8', BAMBOO),
  spacer(100),

  // SECTION 4.9 FOOTER
  h2('4.9 — FOOTER'),
  dividerPara(BAMBOO),
  kvTable([
    ['Background',         '--night (#1A160E)'],
    ['Top border',         '1px solid rgba(200,169,110, 0.1)'],
    ['Layout',             '3-column grid: 2fr 1fr 1fr · gap 3rem · padding 4rem 2.5rem 2rem'],
    ['Bottom bar',         'Full-width · border-top 1px solid rgba(sand,0.06) · flex space-between'],
  ]),
  specBox([
    'COLUMN 1 — Brand',
    '  Logo wordmark: "Aloha Café" · Cormorant · 1.8rem · --sand · "Café" italic · --warm',
    '  Tagline: "Where the city exhales. · Vijayawada" · Jost · 0.8rem · rgba(sand,0.3) · 300',
    '  Social icons row: Instagram · Google Maps · WhatsApp',
    '  Icon style: 36×36px circles · border 1px rgba(sand,0.12) · hover border rgba(bamboo,0.4)',
    '',
    'COLUMN 2 — Navigate',
    '  Title: "Navigate" · 0.62rem Jost ALL CAPS · rgba(sand,0.25)',
    '  Links: Menu · Our Space · Gallery · Reserve',
    '  Link style: Jost · 0.8rem · rgba(sand,0.4) · 300 · hover: --bamboo',
    '',
    'COLUMN 3 — Visit',
    '  Title: "Visit"',
    '  Links: Get Directions · Parking Info · Private Events · Contact Us',
    '',
    'BOTTOM BAR:',
    '  Left:  "© 2024 Aloha Café · Vijayawada · All rights reserved"',
    '  Right: "Designed with care · Built for the backyard"',
    '  Style: Jost · 0.65rem · rgba(sand,0.2) · letter-spacing 0.1em',
  ], 'F0EAD8', BAMBOO),
];

// ─── SECTION 5: BUTTONS ──────────────────────────────────────────────────────
const buttonSection = [
  sectionLabel('SECTION 5 — BUTTON SYSTEM'),
  spacer(60),
  h1italic('Three Button Types. ', 'Used With Restraint.'),
  body('Never use more than two buttons in the same visual cluster. Never stack three buttons vertically. Button text is always Jost, 500, ALL CAPS, letter-spacing 0.16em.'),
  spacer(80),
  threeColTable(
    ['Button Type', 'Specification', 'When to Use'],
    [
      ['Primary (Filled)',       'bg: --bamboo · text: --night · pill: border-radius 3rem · padding 0.85rem 2.2rem\nHover: bg --warm · transform translateY(-2px) · transition 0.25s',             'Main CTA: "Find Your Table" · "View Full Menu" · "Confirm on WhatsApp"'],
      ['Primary Dark (on sand)', 'bg: --palm · text: --sand · same pill shape\nHover: bg --palm-mid · transform translateY(-2px)',                                                                        'CTAs on light backgrounds: nav CTA · section CTAs on sand/cream sections'],
      ['Ghost / Text Link',      'No bg · no border · text: --rust or rgba(sand,0.6)\nBottom border only: 1px solid rgba(rust,0.4)\nHover: full opacity text · transition 0.2s',                         '"Our Story →" · "Explore the Space →" · "Open Instagram →"'],
    ]
  ),
  spacer(80),
  specBox([
    'BUTTON RULES:',
    '',
    '1. Primary button: --bamboo (#C8A96E) bg when placed on dark (--night) sections.',
    '   Primary button: --palm (#2B4A2A) bg when placed on light (--sand/--cream) sections.',
    '2. Never use --palm button on a --palm background section.',
    '3. Ghost CTA always appears AFTER the primary CTA, never before.',
    '4. Button font: Jost · 500 · 0.72rem · letter-spacing 0.16em · ALL CAPS. No exceptions.',
    '5. No box-shadow on any button. Transform on hover instead.',
    '6. Disabled state: opacity 0.4 · no transform · cursor not-allowed.',
    '7. Loading state: subtle opacity pulse animation — never a spinner on a premium site.',
  ], 'F0EAD8', BAMBOO),
];

// ─── SECTION 6: ANIMATIONS ───────────────────────────────────────────────────
const animationSection = [
  sectionLabel('SECTION 6 — ANIMATIONS & INTERACTIONS'),
  spacer(60),
  h1italic('Calm Motion. ', 'Never Flashy.'),
  body('All animations serve the content — they should feel like the page breathing, not a software demo. Prefer CSS transitions over JavaScript where possible. No bounce effects. No spring physics. No infinite loops except the scroll indicator.'),
  spacer(80),
  h2('Scroll-Triggered Reveals'),
  body('Use IntersectionObserver to trigger fade-up animations on section entry. One animation class: .reveal. Applied to: section headlines, body paragraphs, cards, the gallery grid.'),
  specBox([
    '.reveal {',
    '  opacity: 0;',
    '  transform: translateY(24px);',
    '  transition: opacity 0.7s ease, transform 0.7s ease;',
    '}',
    '.reveal.visible {',
    '  opacity: 1;',
    '  transform: translateY(0);',
    '}',
    '',
    'IntersectionObserver threshold: 0.15',
    'Apply staggered animation-delay to card grids: 0s, 0.1s, 0.2s per card',
    'Eyebrow reveals first, then headline (0.1s later), then body (0.15s later)',
  ], 'F0EAD8', PALM),
  spacer(80),
  h2('Hover Interactions'),
  kvTable([
    ['Menu cards',        'transform: translateY(-3px) · top border gradient fades in from opacity 0'],
    ['Experience cards',  'border-color: rgba(bamboo, 0.3) · transition 0.25s'],
    ['Gallery cells',     'inner image: transform scale(1.04) · transition 0.4s ease'],
    ['Nav links',         'color → --bamboo · transition 0.2s'],
    ['Social icons',      'border-color → rgba(bamboo,0.4) · transition 0.2s'],
    ['Footer links',      'color → --bamboo · transition 0.2s'],
    ['CTA buttons',       'background shift + translateY(-2px) · transition 0.25s · no shadow'],
    ['Review cards',      'no hover effect — keep them static and readable'],
  ]),
  spacer(80),
  h2('Page Load Sequence (Hero)'),
  body('Stagger in this exact order. Each element uses the same @keyframes fadeUp:'),
  specBox([
    '@keyframes fadeUp {',
    '  from { opacity: 0; transform: translateY(18px); }',
    '  to   { opacity: 1; transform: translateY(0); }',
    '}',
    '',
    '0.3s delay:  Eyebrow label',
    '0.5s delay:  H1 headline',
    '0.7s delay:  Subheadline paragraph',
    '0.9s delay:  CTA button row',
    '1.3s delay:  Badge circle (opacity only, no translateY)',
    '1.5s delay:  Scroll indicator (opacity only)',
    '',
    'All use: animation-fill-mode: both; animation-duration: 0.8–0.9s; ease timing',
  ], 'F0EAD8', PALM),
  spacer(80),
  h2('Scroll Indicator Animation'),
  specBox([
    '@keyframes scrollPulse {',
    '  0%, 100% { opacity: 0.4; transform: scaleY(1); }',
    '  50%       { opacity: 0.8; transform: scaleY(0.85); }',
    '}',
    'Apply to the vertical line · animation: scrollPulse 2s ease infinite',
    'Hide scroll indicator after user scrolls 100px (JS or scroll-timeline CSS)',
  ], 'F0EAD8', PALM),
];

// ─── SECTION 7: MOBILE ───────────────────────────────────────────────────────
const mobileSection = [
  sectionLabel('SECTION 7 — MOBILE RESPONSIVENESS'),
  spacer(60),
  h1italic('Mobile First. ', '60% of Your Visitors.'),
  body('Design mobile-first. The desktop version is an enhancement of the mobile layout, not the other way around. Breakpoints: mobile < 768px · tablet 768–1024px · desktop > 1024px.'),
  spacer(80),
  h2('Navbar — Mobile'),
  kvTable([
    ['Nav links',       'Hidden at < 768px'],
    ['Desktop CTA',     'Hidden at < 768px'],
    ['Hamburger icon',  'Top-right · 3 lines · --bamboo colour · 24×24px'],
    ['Menu panel',      'Full-screen slide-in from right · --night bg · 4 links + WhatsApp CTA button centred'],
    ['Panel animation', 'transform: translateX(100%) → translateX(0) · transition 0.3s ease'],
    ['Close button',    'Top-right × icon · --sand colour'],
  ]),
  spacer(80),
  h2('Hero — Mobile'),
  kvTable([
    ['Image crop',      'Image 2 portrait orientation — crop to show bamboo wall + one globe light'],
    ['Headline size',   'clamp(2.8rem, 8vw, 3.5rem) — larger relative to screen'],
    ['Content padding', '0 1.5rem 3.5rem — left-bottom anchored'],
    ['Button layout',   'Stack vertically · primary button full-width · ghost CTA centred below'],
    ['Badge',           'Hide the circular badge on mobile — too small'],
    ['Bamboo SVG art',  'Hide on mobile — decorative only'],
  ]),
  spacer(80),
  h2('All 2-Column Sections — Mobile'),
  body('Every two-column desktop layout (About, Visit) collapses to single column on mobile. Image or visual always stacks ABOVE the text, not below.'),
  kvTable([
    ['Column gap on mobile', 'Replaced by margin-bottom: 2rem between stacked items'],
    ['Image height on mobile','280px fixed height · object-fit: cover · border-radius 0.75rem'],
    ['Text max-width',       'Full width minus padding — no artificial max-width on mobile'],
  ]),
  spacer(80),
  h2('Card Grids — Mobile'),
  kvTable([
    ['Menu cards (3-col desktop)',     'Single column on mobile'],
    ['Experience cards (3-col desktop)','Single column on mobile'],
    ['Review cards (3-col desktop)',   'Single column on mobile'],
    ['Ethos strip (3-col desktop)',    'Single column · remove column borders · add bottom border instead'],
    ['Gallery mosaic',                 'Horizontal swipe carousel — NOT a grid on mobile (see below)'],
  ]),
  spacer(80),
  h2('Gallery — Mobile'),
  specBox([
    'Replace the asymmetric grid with a horizontal swipe carousel:',
    '',
    '  container: overflow-x: auto; scroll-snap-type: x mandatory; display: flex',
    '  each cell:  width: 88vw; flex-shrink: 0; scroll-snap-align: start',
    '  gap:        0.8rem between cells',
    '  peek:       right edge of next cell visible (the 12vw remainder)',
    '',
    'Order: Image 2 (full-height hero shot) first · then 3 · 4 · 1 crop',
    'Show dot indicators below carousel (CSS only, no JS required)',
    'Smooth scroll behaviour: scroll-behavior: smooth on container',
  ], 'F0EAD8', BAMBOO),
  spacer(80),
  h2('Sticky CTA Bar — Mobile Only'),
  specBox([
    'Display: fixed · bottom: 0 · left: 0 · right: 0 · z-index: 200',
    'Background: --palm (#2B4A2A) · padding: 0.9rem 1.5rem',
    'Content: "📲 Reserve a Table" · Jost · 500 · 0.75rem · --sand · centred',
    'Border-top: 1px solid rgba(sand, 0.1)',
    '',
    'Hide this bar:',
    '  — when the #visit section is in viewport (user is already at the form)',
    '  — when the mobile nav panel is open',
    '',
    'Show this bar: all other scroll positions',
    '',
    'Tap action: smooth-scroll to #visit section · or open WhatsApp directly',
  ], 'F0EAD8', BAMBOO),
  spacer(80),
  h2('Typography Scale — Mobile'),
  kvTable([
    ['Hero H1',        'clamp(2.8rem, 8vw, 3.5rem)'],
    ['Section H2',     '2rem (fixed) — do not clamp below 1.8rem'],
    ['Body copy',      '1rem / line-height 1.8 — slightly larger than desktop for readability'],
    ['Eyebrow',        '0.65rem — same as desktop'],
    ['Card names',     '1.1rem Cormorant'],
    ['CTA buttons',    '0.78rem — slightly larger than desktop for tap targets'],
    ['Minimum tap target','44×44px for all interactive elements (WCAG AA)'],
  ]),
];

// ─── SECTION 8: IMAGES ───────────────────────────────────────────────────────
const imageSection = [
  sectionLabel('SECTION 8 — IMAGE PLACEMENT & STRATEGY'),
  spacer(60),
  h1italic('Four Photos. ', 'Used With Intention.'),
  spacer(80),
  threeColTable(
    ['Image', 'Placement', 'Treatment'],
    [
      ['Image 2 — Bamboo corridor at night',   '★ HERO background (primary). Also: gallery cell 1 (large).',    'Full-bleed. Dark overlay gradient. Slightly desaturated (filter: saturate(0.85)). Darken bottom 60%.'],
      ['Image 3 — Dusk exterior with palm',    'About section (right column). Also: gallery cell 2.',            'Crop to exclude neighbouring building. Slight warm colour grade. object-position: center top.'],
      ['Image 4 — Daytime arrival/entrance',   'Visit section background (low opacity) or gallery cell 3.',      'Desaturate slightly. overlay rgba(night, 0.3). Best at dusk — note this to client.'],
      ['Image 1 — Interior terrace (daytime)', 'Gallery only — cells 4 and 5 as cropped details.',              'NEVER use wide-angle crop. Crop to: (a) pendant lights cluster, or (b) tiki mural closeup only. Apply warm filter.'],
    ]
  ),
  spacer(80),
  specBox([
    'IMAGE TECHNICAL REQUIREMENTS:',
    '',
    '— All images: loading="lazy" except hero (loading="eager")',
    '— All images: object-fit: cover on fixed-height containers',
    '— Hero background: CSS background-image for best performance',
    '— WebP format preferred · JPEG fallback',
    '— Hero image: minimum 1920×1080px source · compress to < 300KB',
    '— Gallery images: minimum 800×600px each',
    '— Alt text: descriptive but concise (e.g. "Bamboo garden corridor at Aloha Café at night")',
    '',
    'IMPORTANT CLIENT NOTE:',
    'Commission a 30-minute golden-hour photo shoot (6–7 PM). A single good',
    'evening shoot of the garden, verandah, and interior lights will dramatically',
    'elevate the site. Priority shots: bamboo corridor close-ups · candid guests',
    'in warm light · overhead pendant shot · coconut palm from below at dusk.',
  ], 'F0EAD8', RUST),
];

// ─── SECTION 9: NO-FOOD-PHOTO STRATEGY ───────────────────────────────────────
const foodPhotoSection = [
  sectionLabel('SECTION 9 — MENU WITHOUT FOOD PHOTOGRAPHY'),
  spacer(60),
  h1italic('Typography as ', 'Sensory Design.'),
  body('This is a deliberate creative choice used by the world\'s best restaurant websites — Noma, Aesop, Ottolenghi. Lead with atmosphere. The food becomes something to discover in person.'),
  spacer(60),
  h2('The Three-Part Strategy'),
  kvTable([
    ['Strategy 1: Menu as Typography',    'Use Cormorant Garamond for item names + poetic 1-line descriptors in Jost. The copy itself is the sensory experience. "Slow-dripped overnight, ice-settled, unhurried" sells the drink.'],
    ['Strategy 2: Ingredient/Texture UGC','Prompt guests on Instagram to photograph their food. Embed the hashtag feed. Real food in your actual space > any styled shoot on a white table.'],
    ['Strategy 3: Texture Details',       'If any photography is available: shoot textures, not plates. Coffee being poured. Ice in a glass. A half-eaten dessert on the bamboo table. 5 seconds of authentic beats any styled shot.'],
  ]),
  spacer(80),
  h2('Menu Copy Guidelines for the Builder'),
  body('When the AI builder generates menu item descriptions, use this formula:'),
  specBox([
    'MENU ITEM COPY FORMULA:',
    '',
    'Name:        [Simple, evocative] — e.g. "Cold Brew Ritual", "Garden Lemonade"',
    '             NOT: "Special House-Made Cold Brew Coffee with Ice"',
    '',
    'Descriptor:  [1 line · sensory · no adjective overload]',
    '             "Slow-dripped overnight, ice-settled, unhurried."',
    '             NOT: "Our amazing cold brew is made fresh daily with the finest beans!"',
    '',
    'Price:       [Bare number in Cormorant serif · no "₹" with a rupee word after]',
    '             "₹180" — not "Rs. 180/-"',
    '',
    'Badge:       [Optional · 1 word max] "Bestseller" "New" "Chef\'s Pick" "Seasonal"',
    '             Not: "Fan Favourite!!!" "Must Try!"',
  ], 'F0EAD8', PALM),
];

// ─── SECTION 10: GRETA PROMPT ────────────────────────────────────────────────
const gretaPromptSection = [
  sectionLabel('SECTION 10 — READY-TO-USE GRETA.SH PROMPT'),
  spacer(60),
  h1italic('The Master Prompt. ', 'Copy and Paste.'),
  body('Use the following prompt verbatim in the Greta.sh builder. It consolidates all design decisions into a single generation instruction.'),
  spacer(80),
  darkBox([
    'GRETA.SH WEBSITE GENERATION PROMPT',
    '════════════════════════════════════════════════════════════',
    '',
    'Build a premium, cinematic, tropical café website for "Aloha Café — The Back Yard" located in Vijayawada, India.',
    '',
    'DESIGN IDENTITY:',
    '  Style:    Boutique tropical lounge. NOT a generic restaurant template.',
    '  Mood:     Cinematic, calm, warm, premium, editorial.',
    '  Reference: Aesop + Soho House + tropical garden hotel editorial sites.',
    '',
    'TYPOGRAPHY (load from Google Fonts — exact names required):',
    '  Display: Cormorant Garamond — weight 300 + 300 italic only',
    '  UI/Body: Jost — weight 300, 400, 500',
    '  No other fonts. No system fonts.',
    '',
    'COLOUR PALETTE (exact hex — no substitutions):',
    '  --night:      #1A160E  (backgrounds: hero, footer, experience section)',
    '  --palm:       #2B4A2A  (CTA buttons on light sections)',
    '  --bamboo:     #C8A96E  (headlines, highlights, primary CTA on dark)',
    '  --warm:       #E8B86A  (italic serif accents ONLY)',
    '  --sand:       #F4EEE2  (text on dark, card fills)',
    '  --cream:      #FAF7F2  (primary page background)',
    '  --sand2:      #EDE4D3  (borders, dividers)',
    '  --bamboo-d:   #8B6E3A  (italic headline text)',
    '  --muted:      #8A7A68  (body copy, metadata)',
    '',
    'HOMEPAGE SECTIONS (build in this exact order):',
    '  1. Fixed navbar (dark, frosted glass, bamboo-lit)',
    '  2. Hero (full-viewport, Image 2 background, bottom-left content)',
    '  3. Ethos strip (3-col, sand bg, numbered moods)',
    '  4. About / Our Space (2-col, cream bg, Image 3 right)',
    '  5. Menu showcase (sand bg, typographic cards, no food photos)',
    '  6. Experience — Why Aloha (3-col, night bg dark section)',
    '  7. Gallery mosaic (asymmetric grid, cream bg)',
    '  8. Reviews (3-col cards, sand bg)',
    '  9. Visit & Reserve (2-col, palm bg, WhatsApp form)',
    '  10. Footer (3-col, night bg, minimal)',
    '',
    'HERO COPY (exact):',
    '  Eyebrow:     "Aloha — The Back Yard · Vijayawada"',
    '  Headline:    "Where the city exhales."',
    '               ("exhales." = italic, --warm colour, own line)',
    '  Subtext:     "A garden café tucked behind bamboo walls and swaying palms.',
    '               Open-air seating, warm evenings, and a menu worth lingering over."',
    '  Primary CTA: "Find Your Table"',
    '  Ghost CTA:   "Explore the Space →"',
    '',
    'NAVIGATION:',
    '  Logo:       "Aloha" (Cormorant 300, --sand) + "Café" (italic, --warm)',
    '  Links:      Menu · Our Space · Gallery · Visit',
    '  CTA:        "Reserve a Table" (pill, --palm bg, --sand text)',
    '  Behaviour:  Fixed position, dark frosted glass, scroll-triggered opacity increase',
    '',
    'ANIMATIONS:',
    '  Hero load:  Staggered fadeUp (opacity+translateY) with 0.2s delays between elements',
    '  Scroll:     IntersectionObserver reveal on all section content',
    '  Hover:      Subtle translateY on cards, colour shifts on links — nothing bouncy',
    '',
    'MOBILE:',
    '  Mobile-first layout. Hamburger nav. Sticky bottom WhatsApp CTA.',
    '  Gallery converts to horizontal swipe carousel on mobile.',
    '  All 2-col sections stack vertically. Image above text always.',
    '',
    'TONE — DO NOT:',
    '  × Use exclamation marks anywhere',
    '  × Use the words: amazing, delicious, fresh, cosy, vibe (in body copy)',
    '  × Use stock food photography',
    '  × Add more than 3 navigation items',
    '  × Use bright colours, neon, gradients with multiple hues',
    '  × Use drop shadows on cards or buttons',
    '',
    'TONE — DO:',
    '  ✓ Use short, confident, poetic copy throughout',
    '  ✓ Let white space do the work',
    '  ✓ Italicise key emotional words in headlines',
    '  ✓ Lead with atmosphere, then amenities, then menu',
    '  ✓ Make the WhatsApp reservation CTA the primary conversion action',
    '════════════════════════════════════════════════════════════',
  ], '1A160E', BAMBOO),
];

// ─── ASSEMBLE DOCUMENT ────────────────────────────────────────────────────────
const doc = new Document({
  styles: {
    default: {
      document: { run: { font: 'Jost', size: 20, color: NIGHT } }
    },
    paragraphStyles: [
      {
        id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run:       { font: 'Cormorant Garamond', size: 52, color: NIGHT, bold: false },
        paragraph: { spacing: { before: 200, after: 160 }, outlineLevel: 0 }
      },
      {
        id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run:       { font: 'Cormorant Garamond', size: 40, color: NIGHT, bold: false },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 1 }
      },
      {
        id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run:       { font: 'Cormorant Garamond', size: 28, color: NIGHT, bold: false },
        paragraph: { spacing: { before: 160, after: 100 }, outlineLevel: 2 }
      },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
      }
    },
    children: [
      ...coverSection,
      ...brandSection,
      ...navSection,
      ...homepageSection,
      ...buttonSection,
      ...animationSection,
      ...mobileSection,
      ...imageSection,
      ...foodPhotoSection,
      ...gretaPromptSection,
      spacer(200),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/mnt/user-data/outputs/AlohaCafe_GretaSpec.docx', buf);
  console.log('SUCCESS: AlohaCafe_GretaSpec.docx written');
});
