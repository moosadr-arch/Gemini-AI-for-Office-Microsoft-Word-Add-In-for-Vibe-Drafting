import assert from 'assert';
import { insertContentAsNativeParagraphs } from '../src/taskpane/modules/docx-redline-js-integration/word-redline-runner.js';

// insertContentAsNativeParagraphs drives native Word APIs (insertText /
// insertParagraph) which rely on the document's change-tracking mode to redline.
// We can't run Word in Node, but we can verify the logic with a mock paragraph:
// which line goes where, and Replace-vs-After placement.

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

const mockContext = { sync: async () => {} };

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

await testFillAnchorMode();
await testAppendMode();
await testSingleLine();

console.log('empty_paragraph_insertion_tests passed');
