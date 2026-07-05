import './setup-xml-provider.mjs';
import assert from 'assert';
import {
  applyRedlineChangesToWordContext,
  insertContentAsNativeParagraphs,
  prepareNativeMarkdownParagraphs,
  segmentNativeInsertionBlocks
} from '../src/taskpane/modules/docx-redline-js-integration/word-redline-runner.js';

// insertContentAsNativeParagraphs drives native Word APIs (insertText /
// insertParagraph) which rely on the document's change-tracking mode to redline.
// Bullet-list and markdown-table blocks are rendered through the reconciliation
// engine's OOXML generators and inserted with insertOoxml on a placeholder
// paragraph. We can't run Word in Node, but we can verify the logic with mock
// paragraphs: which line goes where, Replace-vs-After placement, and what OOXML
// each placeholder receives.

if (!globalThis.Word) {
  globalThis.Word = { ChangeTrackingMode: { off: 'Off' } };
}

function makeMockParagraph(log, id) {
  return {
    id,
    insertText(text, location) {
      log.push({ op: 'insertText', id, text, location });
    },
    insertParagraph(text, location) {
      const childId = `${id}>after`;
      log.push({ op: 'insertParagraph', id, text, location });
      return makeMockParagraph(log, childId);
    }
  };
}

function makeFormattingMockParagraph(log, id) {
  const paragraph = makeMockParagraph(log, id);
  paragraph.font = {};
  paragraph.getRange = () => ({
    search(text, options) {
      const targetRange = { font: {} };
      log.push({ op: 'search', id, text, options, targetRange });
      return {
        items: [targetRange],
        load(property) {
          log.push({ op: 'loadSearch', id, property });
        }
      };
    }
  });
  return paragraph;
}

// OOXML-capable mock: also records insertOoxml calls (list/table block path).
// Child ids are unique so placement can be asserted precisely.
function makeOoxmlMockParagraph(log, id, opts = {}) {
  let childCount = 0;
  return {
    id,
    insertText(text, location) {
      log.push({ op: 'insertText', id, text, location });
    },
    insertParagraph(text, location) {
      childCount += 1;
      const child = makeOoxmlMockParagraph(log, `${id}>c${childCount}`, opts);
      log.push({ op: 'insertParagraph', id, childId: child.id, text, location });
      return child;
    },
    insertOoxml(ooxml, location) {
      log.push({ op: 'insertOoxml', id, ooxml, location });
      if (opts.failOoxml) {
        throw new Error('mock insertOoxml failure');
      }
    }
  };
}

const mockContext = { sync: async () => {} };
const ooxmlMockContext = {
  sync: async () => {},
  document: {
    load() {},
    changeTrackingMode: 'Off'
  }
};

function insertedTexts(log) {
  return log
    .filter((e) => e.op === 'insertText' || e.op === 'insertParagraph')
    .map((e) => e.text);
}

async function testFillAnchorMode() {
  // Empty-paragraph insertion: first line replaces the anchor, rest go after.
  const log = [];
  const anchor = makeMockParagraph(log, 'P5');
  await insertContentAsNativeParagraphs(mockContext, anchor, 'Line one\nLine two\nLine three', { fillAnchor: true });

  assert.deepStrictEqual(log[0], { op: 'insertText', id: 'P5', text: 'Line one', location: 'Replace' });
  // Remaining lines are inserted After, each chained off the previous insertion.
  assert.strictEqual(log[1].op, 'insertParagraph');
  assert.strictEqual(log[1].text, 'Line two');
  assert.strictEqual(log[1].location, 'After');
  assert.strictEqual(log[2].text, 'Line three');
  assert.strictEqual(log.length, 3);
}

async function testAppendMode() {
  // End-of-document append: ALL lines inserted after the anchor (last paragraph).
  const log = [];
  const last = makeMockParagraph(log, 'Plast');
  await insertContentAsNativeParagraphs(mockContext, last, 'First new\nSecond new');

  assert.strictEqual(log.length, 2);
  assert.deepStrictEqual(
    log.map((e) => [e.op, e.text, e.location]),
    [['insertParagraph', 'First new', 'After'], ['insertParagraph', 'Second new', 'After']]
  );
  // No insertText (nothing is replaced when appending).
  assert.ok(!log.some((e) => e.op === 'insertText'));
}

async function testSingleLine() {
  const log = [];
  const anchor = makeMockParagraph(log, 'P5');
  await insertContentAsNativeParagraphs(mockContext, anchor, 'Only line', { fillAnchor: true });
  assert.deepStrictEqual(log, [{ op: 'insertText', id: 'P5', text: 'Only line', location: 'Replace' }]);
}

async function testNativeInsertionAppliesMarkdownFormatting() {
  const log = [];
  const anchor = makeFormattingMockParagraph(log, 'P5');
  await insertContentAsNativeParagraphs(
    mockContext,
    anchor,
    'This Agreement is between **Electronic Arts Inc.** and **[Company Name]**.',
    { fillAnchor: true }
  );

  assert.deepStrictEqual(log[0], {
    op: 'insertText',
    id: 'P5',
    text: 'This Agreement is between Electronic Arts Inc. and [Company Name].',
    location: 'Replace'
  });
  const searches = log.filter((entry) => entry.op === 'search');
  assert.deepStrictEqual(
    searches.map((entry) => entry.text),
    ['Electronic Arts Inc.', '[Company Name]']
  );
  assert.strictEqual(searches[0].targetRange.font.bold, true);
  assert.strictEqual(searches[1].targetRange.font.bold, true);
}

function testNativeMarkdownPreparation() {
  const prepared = prepareNativeMarkdownParagraphs([
    '# MUTUAL NON-DISCLOSURE AGREEMENT',
    'This Agreement is between **Electronic Arts Inc.** and **[Company Name]**.',
    '**1. Purpose.** The Parties wish to explore a potential business relationship.'
  ].join('\n'));

  assert.strictEqual(prepared.length, 3);
  assert.strictEqual(prepared[0].text, 'MUTUAL NON-DISCLOSURE AGREEMENT');
  assert.strictEqual(prepared[0].headingLevel, 1);
  assert.ok(prepared[0].formatHints.some((hint) => hint.format.bold && hint.start === 0));

  assert.strictEqual(
    prepared[1].text,
    'This Agreement is between Electronic Arts Inc. and [Company Name].'
  );
  assert.strictEqual(prepared[1].formatHints.length, 2);
  assert.ok(!prepared[1].text.includes('**'), 'bold delimiters should not leak into inserted text');

  assert.strictEqual(
    prepared[2].text,
    '1. Purpose. The Parties wish to explore a potential business relationship.'
  );
  assert.ok(!prepared[2].text.includes('**'), 'section-heading delimiters should be stripped');
}

function testSegmentation() {
  // Bullets between prose; the blank line after the list is consumed (the
  // sentinel/spacing paragraph inserted after the block represents it).
  assert.deepStrictEqual(
    segmentNativeInsertionBlocks('Intro\n* a\n* b\n\nOutro'),
    [
      { kind: 'text', lines: ['Intro'] },
      { kind: 'list', lines: ['* a', '* b'] },
      { kind: 'text', lines: ['Outro'] }
    ]
  );

  // Bold delimiters and horizontal rules are not bullet markers.
  assert.deepStrictEqual(
    segmentNativeInsertionBlocks('**1. Purpose** text\n---'),
    [{ kind: 'text', lines: ['**1. Purpose** text', '---'] }]
  );

  // A markdown table needs a header row followed by a separator row.
  assert.deepStrictEqual(
    segmentNativeInsertionBlocks('| H1 | H2 |\n|---|---|\n| a | b |'),
    [{ kind: 'table', lines: ['| H1 | H2 |', '|---|---|', '| a | b |'] }]
  );

  // Pipe lines without a separator row stay literal text.
  assert.strictEqual(segmentNativeInsertionBlocks('| a | b |\n| c | d |')[0].kind, 'text');

  // Indented nested markers continue a bullet block.
  assert.deepStrictEqual(
    segmentNativeInsertionBlocks('* a\n  1. sub\n* b'),
    [{ kind: 'list', lines: ['* a', '  1. sub', '* b'] }]
  );

  // Top-level numbered lines stay text: engine-generated lists restart at 1,
  // which would renumber explicit section numbering like "3. Exclusions".
  assert.deepStrictEqual(
    segmentNativeInsertionBlocks('1. Purpose\n2. Scope'),
    [{ kind: 'text', lines: ['1. Purpose', '2. Scope'] }]
  );

  // Bullet variants - and + also start list blocks.
  assert.strictEqual(segmentNativeInsertionBlocks('- a\n+ b')[0].kind, 'list');
}

async function testBulletBlockRendersNativeListOoxml() {
  const log = [];
  const last = makeOoxmlMockParagraph(log, 'Plast');
  await insertContentAsNativeParagraphs(
    ooxmlMockContext,
    last,
    'Intro line\n* First bullet item\n* Second bullet item\n\nOutro line'
  );

  const ooxmlOps = log.filter((e) => e.op === 'insertOoxml');
  assert.strictEqual(ooxmlOps.length, 1, 'the bullet block should be one OOXML insertion');
  const pkg = ooxmlOps[0].ooxml;
  assert.strictEqual(ooxmlOps[0].location, 'Replace');
  assert.ok(pkg.includes('<w:numPr>'), 'list OOXML should bind Word numbering');
  assert.ok(pkg.includes('numbering.xml'), 'package should embed numbering definitions');
  assert.ok(pkg.includes('First bullet item'), 'first item text missing from OOXML');
  assert.ok(pkg.includes('Second bullet item'), 'second item text missing from OOXML');
  assert.ok(pkg.includes('<w:ins'), 'list insertion should carry redlines by default');
  assert.ok(!pkg.includes('* First'), 'raw bullet markers must not leak into OOXML text');

  // No natively inserted line may carry a raw markdown bullet marker.
  for (const text of insertedTexts(log)) {
    assert.ok(!/^\s*[*+-]\s/.test(text || ''), `raw marker leaked into native text: "${text}"`);
  }

  // Placement: intro after anchor, then placeholder + sentinel, OOXML replaces
  // the placeholder, and the outro continues after the sentinel.
  const paraOps = log.filter((e) => e.op === 'insertParagraph');
  assert.deepStrictEqual(paraOps.map((e) => e.text), ['Intro line', '', '', 'Outro line']);
  const placeholderId = paraOps[1].childId;
  const sentinelId = paraOps[2].childId;
  assert.strictEqual(ooxmlOps[0].id, placeholderId, 'OOXML should replace the placeholder paragraph');
  assert.strictEqual(paraOps[3].id, sentinelId, 'text after the list should chain off the sentinel');
}

async function testFillAnchorLeadingBulletsReplaceAnchor() {
  // Whole content is a bullet list and the target paragraph is empty:
  // the list OOXML replaces the anchor paragraph itself.
  const log = [];
  const anchor = makeOoxmlMockParagraph(log, 'P5');
  await insertContentAsNativeParagraphs(ooxmlMockContext, anchor, '* Alpha\n* Beta', { fillAnchor: true });

  const ooxmlOps = log.filter((e) => e.op === 'insertOoxml');
  assert.strictEqual(ooxmlOps.length, 1);
  assert.strictEqual(ooxmlOps[0].id, 'P5', 'list should replace the empty anchor paragraph');
  assert.strictEqual(ooxmlOps[0].location, 'Replace');
  assert.ok(!log.some((e) => e.op === 'insertText'), 'no literal text should be inserted');
}

async function testTableBlockRendersTableOoxml() {
  const log = [];
  const last = makeOoxmlMockParagraph(log, 'Plast');
  await insertContentAsNativeParagraphs(
    ooxmlMockContext,
    last,
    'Parties:\n| Disclosing Party | Receiving Party |\n|---|---|\n| EA | Counterparty |'
  );

  const ooxmlOps = log.filter((e) => e.op === 'insertOoxml');
  assert.strictEqual(ooxmlOps.length, 1, 'the table block should be one OOXML insertion');
  assert.ok(ooxmlOps[0].ooxml.includes('<w:tbl'), 'table OOXML should contain a Word table');
  assert.ok(ooxmlOps[0].ooxml.includes('Disclosing Party'), 'header text missing from table OOXML');
  for (const text of insertedTexts(log)) {
    assert.ok(!(text || '').includes('|'), `raw table row leaked into native text: "${text}"`);
  }
}

async function testOoxmlFailureFallsBackToLiteralBullets() {
  // If Word rejects the OOXML package, degrade to literal bullet characters -
  // never raw markdown markers.
  const log = [];
  const anchor = makeOoxmlMockParagraph(log, 'P5', { failOoxml: true });
  await insertContentAsNativeParagraphs(ooxmlMockContext, anchor, '* Alpha\n* Beta', { fillAnchor: true });

  const texts = insertedTexts(log);
  assert.ok(texts.includes('• Alpha'), 'fallback should render a literal bullet character');
  assert.ok(texts.includes('• Beta'), 'fallback should render every bullet line');
  assert.ok(!texts.some((t) => /^\s*[*+-]\s/.test(t || '')), 'raw markers must not leak in fallback');
}

async function testTopLevelNumberedLinesStayLiteral() {
  // Engine-generated numbered lists restart at 1, so explicit section numbers
  // must be inserted as literal text to avoid renumbering "2." into "1.".
  const log = [];
  const anchor = makeOoxmlMockParagraph(log, 'P5');
  await insertContentAsNativeParagraphs(
    ooxmlMockContext,
    anchor,
    '1. Purpose\nThe Parties wish to explore.\n2. Confidential Information',
    { fillAnchor: true }
  );

  assert.ok(!log.some((e) => e.op === 'insertOoxml'), 'numbered section headings must stay native text');
  assert.deepStrictEqual(insertedTexts(log), [
    '1. Purpose',
    'The Parties wish to explore.',
    '2. Confidential Information'
  ]);
}

async function testApplyRedlineChangesInsertsIntoEmptyParagraphBeforeConversion() {
  const log = [];
  const infos = [];
  const warnings = [];
  const anchor = makeMockParagraph(log, 'P1');
  anchor.text = '';

  const paragraphs = {
    items: [anchor],
    load(property) {
      log.push({ op: 'loadParagraphs', property });
    }
  };
  const context = {
    sync: async () => {},
    document: {
      body: {
        paragraphs
      }
    }
  };

  const result = await applyRedlineChangesToWordContext(
    context,
    [{
      paragraphIndex: 1,
      operation: 'edit_paragraph',
      newContent: 'Silent lines of code,\nContracts bind our virtual worlds,\nKabam stands secure.'
    }],
    {
      onInfo: (message) => infos.push(message),
      onWarn: (message) => warnings.push(message)
    }
  );

  assert.strictEqual(result.changesApplied, 1);
  assert.deepStrictEqual(result.skipped, []);
  assert.ok(infos.some((message) => message.includes('before redline conversion')));
  assert.ok(!warnings.some((message) => message.includes('Target paragraph text is empty')));
  assert.deepStrictEqual(
    log.filter((entry) => entry.op === 'insertText' || entry.op === 'insertParagraph').map((entry) => [entry.op, entry.text, entry.location]),
    [
      ['insertText', 'Silent lines of code,', 'Replace'],
      ['insertParagraph', 'Contracts bind our virtual worlds,', 'After'],
      ['insertParagraph', 'Kabam stands secure.', 'After']
    ]
  );
}

await testFillAnchorMode();
await testAppendMode();
await testSingleLine();
await testNativeInsertionAppliesMarkdownFormatting();
testNativeMarkdownPreparation();
testSegmentation();
await testBulletBlockRendersNativeListOoxml();
await testFillAnchorLeadingBulletsReplaceAnchor();
await testTableBlockRendersTableOoxml();
await testOoxmlFailureFallsBackToLiteralBullets();
await testTopLevelNumberedLinesStayLiteral();
await testApplyRedlineChangesInsertsIntoEmptyParagraphBeforeConversion();

console.log('empty_paragraph_insertion_tests passed');
