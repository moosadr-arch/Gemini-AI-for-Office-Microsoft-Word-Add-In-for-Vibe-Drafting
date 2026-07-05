import assert from 'assert';
import {
  formatAutoCheckpointLabel,
  makeCheckpointRecord,
  idsToEvict,
  migrateLegacyCheckpoints,
  MAX_CHECKPOINTS
} from '../src/taskpane/modules/storage/checkpoint-store.js';

function testFormatAutoCheckpointLabel() {
  const d = new Date('2026-06-12T15:04:05.000Z');
  assert.strictEqual(formatAutoCheckpointLabel('apply_redlines', d), 'auto:apply_redlines:2026-06-12T15:04:05.000Z');
  // Accepts epoch millis / string dates too.
  assert.strictEqual(formatAutoCheckpointLabel('edit_table', d.getTime()), 'auto:edit_table:2026-06-12T15:04:05.000Z');
  // Missing tool name falls back to "unknown".
  assert.ok(formatAutoCheckpointLabel(null, d).startsWith('auto:unknown:'));
}

function testMakeCheckpointRecord() {
  const rec = makeCheckpointRecord('manual', 'PK…', 1234);
  assert.deepStrictEqual(rec, { timestamp: 1234, label: 'manual', ooxml: 'PK…' });
  // Defaults: empty label -> "manual", null ooxml -> "".
  const rec2 = makeCheckpointRecord('', null, 1);
  assert.strictEqual(rec2.label, 'manual');
  assert.strictEqual(rec2.ooxml, '');
}

function testIdsToEvictUnderCap() {
  const records = [{ id: 1 }, { id: 2 }, { id: 3 }];
  assert.deepStrictEqual(idsToEvict(records, 10), []);
  assert.deepStrictEqual(idsToEvict([], 10), []);
  assert.deepStrictEqual(idsToEvict(null, 10), []);
}

function testIdsToEvictOverCap() {
  // 12 records, cap 10 -> evict the 2 oldest by id.
  const records = [];
  for (let i = 1; i <= 12; i++) records.push({ id: i });
  assert.deepStrictEqual(idsToEvict(records, 10), [1, 2]);
  // Unsorted input is handled (sorts by id first).
  const shuffled = [{ id: 5 }, { id: 1 }, { id: 3 }, { id: 2 }, { id: 4 }];
  assert.deepStrictEqual(idsToEvict(shuffled, 3), [1, 2]);
}

function testDefaultCapIsTen() {
  assert.strictEqual(MAX_CHECKPOINTS, 10);
  const records = [];
  for (let i = 1; i <= 11; i++) records.push({ id: i });
  assert.deepStrictEqual(idsToEvict(records), [1]); // uses default MAX_CHECKPOINTS
}

function testMigrateLegacyCheckpoints() {
  const legacy = ['<ooxml1>', '', '<ooxml2>', null, 42, '<ooxml3>'];
  const records = migrateLegacyCheckpoints(legacy, 1000);
  // Only the 3 non-empty strings survive, labeled "migrated", with increasing timestamps.
  assert.strictEqual(records.length, 3);
  assert.deepStrictEqual(records.map((r) => r.ooxml), ['<ooxml1>', '<ooxml2>', '<ooxml3>']);
  assert.ok(records.every((r) => r.label === 'migrated'));
  assert.deepStrictEqual(records.map((r) => r.timestamp), [1000, 1001, 1002]);
  // Non-array input is safe.
  assert.deepStrictEqual(migrateLegacyCheckpoints(null), []);
}

testFormatAutoCheckpointLabel();
testMakeCheckpointRecord();
testIdsToEvictUnderCap();
testIdsToEvictOverCap();
testDefaultCapIsTen();
testMigrateLegacyCheckpoints();

console.log('checkpoint_store_tests passed');
