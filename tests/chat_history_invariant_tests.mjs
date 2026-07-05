import assert from 'assert';
import {
  appendFunctionExchange,
  validateHistoryPairs
} from '../src/taskpane/modules/chat/chat-history.js';

// Helpers to build well-formed turns.
function modelTurn(callNames, extraParts = []) {
  const parts = callNames.map((name) => ({ functionCall: { name, args: {} } }));
  return { role: 'model', parts: [...extraParts, ...parts] };
}
function userTurn(responseNames) {
  return {
    role: 'user',
    parts: responseNames.map((name) => ({
      functionResponse: { name, response: { name, content: [{ text: 'ok' }] } }
    }))
  };
}

function testValidPairAppendsBoth() {
  const history = [{ role: 'user', parts: [{ text: 'hi' }] }];
  const result = appendFunctionExchange(history, modelTurn(['apply_redlines']), userTurn(['apply_redlines']));
  assert.strictEqual(result, history, 'returns the same array');
  assert.strictEqual(history.length, 3);
  assert.strictEqual(history[1].role, 'model');
  assert.strictEqual(history[2].role, 'user');
}

function testMixedModelPartsAllowed() {
  // A model turn may include thought/text parts alongside the functionCall.
  const history = [];
  const mTurn = modelTurn(['edit_list'], [{ text: 'Let me fix that.', thought: true }]);
  appendFunctionExchange(history, mTurn, userTurn(['edit_list']));
  assert.strictEqual(history.length, 2);
}

function testMultipleToolsMatched() {
  const history = [];
  appendFunctionExchange(
    history,
    modelTurn(['apply_redlines', 'insert_comment']),
    userTurn(['apply_redlines', 'insert_comment'])
  );
  assert.strictEqual(history.length, 2);
}

function testMismatchedNameThrowsAndLeavesHistoryUnchanged() {
  const history = [{ role: 'user', parts: [{ text: 'hi' }] }];
  assert.throws(
    () => appendFunctionExchange(history, modelTurn(['apply_redlines']), userTurn(['insert_comment'])),
    /mismatch|uncalled/
  );
  assert.strictEqual(history.length, 1, 'history must be untouched on throw');
}

function testCountMismatchThrows() {
  const history = [];
  // 2 calls to the same tool, only 1 response.
  assert.throws(
    () => appendFunctionExchange(history, modelTurn(['apply_redlines', 'apply_redlines']), userTurn(['apply_redlines'])),
    /mismatch/
  );
  assert.strictEqual(history.length, 0);
}

function testNoFunctionCallThrows() {
  const history = [];
  assert.throws(
    () => appendFunctionExchange(history, { role: 'model', parts: [{ text: 'hello' }] }, userTurn([])),
    /no functionCall/
  );
  assert.strictEqual(history.length, 0);
}

function testBadShapesThrow() {
  assert.throws(() => appendFunctionExchange(null, modelTurn(['x']), userTurn(['x'])), /history must be an array/);
  assert.throws(() => appendFunctionExchange([], { role: 'user', parts: [] }, userTurn([])), /modelTurn must be/);
  assert.throws(() => appendFunctionExchange([], modelTurn(['x']), { role: 'model', parts: [] }), /userTurn must be/);
}

function testValidateHistoryPairsLeavesBuiltHistoryUnchanged() {
  // A history assembled solely via appendFunctionExchange must survive validation.
  let history = [{ role: 'user', parts: [{ text: 'do two things' }] }];
  appendFunctionExchange(history, modelTurn(['apply_redlines']), userTurn(['apply_redlines']));
  appendFunctionExchange(history, modelTurn(['insert_comment']), userTurn(['insert_comment']));
  const validated = validateHistoryPairs(history);
  assert.deepStrictEqual(validated, history, 'validation should not drop any turns');
}

testValidPairAppendsBoth();
testMixedModelPartsAllowed();
testMultipleToolsMatched();
testMismatchedNameThrowsAndLeavesHistoryUnchanged();
testCountMismatchThrows();
testNoFunctionCallThrows();
testBadShapesThrow();
testValidateHistoryPairsLeavesBuiltHistoryUnchanged();

console.log('chat_history_invariant_tests passed');
