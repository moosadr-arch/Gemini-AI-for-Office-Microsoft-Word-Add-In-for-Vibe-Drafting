import { applyWordOperation } from './word-operation-runner.js';
import { toScopedSharedRedlineOperation } from '@ansonlai/docx-redline-js/orchestration/redline-operation-converter.js';

/**
 * Insert `content` (\n = paragraph break) as native, tracked Word paragraphs.
 *
 * Used for inserting into an empty paragraph or appending at end-of-document —
 * cases the OOXML reconciliation engine cannot handle (Word rejects the generated
 * package with InvalidArgument). Native insertText/insertParagraph rely on the
 * document's change-tracking mode (already set to trackAll by setChangeTrackingForAi
 * when redlines are enabled), so insertions are tracked automatically and revert
 * cleanly. Plain text only — markdown formatting is not applied here.
 *
 * @param {boolean} [opts.fillAnchor] - replace the anchor paragraph's (empty) text
 *   with the first line, then insert the rest after it. When false, all lines are
 *   inserted after the anchor (append).
 */
export async function insertContentAsNativeParagraphs(context, anchorParagraph, content, opts = {}) {
    const lines = String(content == null ? '' : content).split('\n');
    if (lines.length === 0) return false;

    let anchor = anchorParagraph;
    let startIdx = 0;
    if (opts.fillAnchor) {
        anchorParagraph.insertText(lines[0], 'Replace');
        await context.sync();
        startIdx = 1;
    }
    for (let i = startIdx; i < lines.length; i += 1) {
        anchor = anchor.insertParagraph(lines[i], 'After');
        await context.sync();
    }
    return true;
}
import { isMarkdownTableText } from '@ansonlai/docx-redline-js/core/paragraph-targeting.js';
import { isLikelyStructuredTableSourceParagraph } from '@ansonlai/docx-redline-js/core/table-targeting.js';

function normalizeNeedleText(value) {
    if (value == null) return '';
    return String(value)
        .replace(/\\n/g, '\n')
        .replace(/\\t/g, '\t')
        .replace(/\\r/g, '\r')
        .trim();
}

function normalizeReplacementText(value) {
    if (value == null) return '';
    return String(value)
        .replace(/\\n/g, '\n')
        .replace(/\\t/g, '\t')
        .replace(/\\r/g, '\r');
}

function getReplacementText(change, operationName) {
    if (operationName === 'edit_paragraph') {
        return change?.newContent;
    }
    if (operationName === 'replace_paragraph' || operationName === 'replace_range') {
        return change?.content ?? change?.newContent ?? change?.replacementText;
    }
    if (operationName === 'modify_text') {
        return change?.replacementText;
    }
    return null;
}

function hasStructuralReplacementField(change) {
    if (!change || typeof change !== 'object') return false;
    return Object.prototype.hasOwnProperty.call(change, 'content')
        || Object.prototype.hasOwnProperty.call(change, 'newContent');
}

function setReplacementText(change, operationName, replacementText) {
    if (operationName === 'edit_paragraph') {
        change.newContent = replacementText;
        return;
    }
    if (operationName === 'replace_paragraph' || operationName === 'replace_range') {
        change.content = replacementText;
    }
}

function splitPipeCells(line) {
    return String(line || '')
        .split('|')
        .map(cell => cell.trim())
        .filter(cell => cell.length > 0);
}

function stripTrailingColon(text) {
    return String(text || '').trim().replace(/:\s*$/, '');
}

function markdownRow(cells) {
    return `| ${cells.map(cell => String(cell || '').trim()).join(' | ')} |`;
}

function markdownSeparator(columnCount) {
    return markdownRow(new Array(Math.max(1, columnCount)).fill('---'));
}

function markdownTableForLabeledColumns(columns) {
    const safeColumns = Array.isArray(columns)
        ? columns.filter(column => column && column.label)
        : [];
    if (safeColumns.length < 2) return null;

    const maxDetailRows = safeColumns.reduce(
        (max, column) => Math.max(max, Array.isArray(column.details) ? column.details.length : 0),
        0
    );

    const bodyRows = [];
    for (let rowIndex = 0; rowIndex < Math.max(1, maxDetailRows); rowIndex += 1) {
        bodyRows.push(markdownRow(
            safeColumns.map(column => String(column.details?.[rowIndex] || '').trim())
        ));
    }

    return [
        markdownRow(safeColumns.map(column => column.label)),
        markdownSeparator(safeColumns.length),
        ...bodyRows
    ].join('\n');
}

function isMarkdownSeparatorRow(line) {
    const normalized = String(line || '').replace(/\s+/g, '');
    return /^\|:?-{3,}:?(\|:?-{3,}:?)+\|?$/.test(normalized);
}

function splitCellBreaks(cell) {
    return String(cell || '')
        .split(/<br\s*\/?>/gi)
        .map(part => part.trim());
}

function expandMarkdownTableHtmlBreaks(text) {
    const lines = normalizeReplacementText(text)
        .trim()
        .split(/\r?\n/)
        .map(line => line.trim())
        .filter(line => line.startsWith('|'));
    if (lines.length === 0 || !lines.some(line => /<br\s*\/?>/i.test(line))) {
        return normalizeReplacementText(text).trim();
    }

    const expanded = [];
    for (const line of lines) {
        if (isMarkdownSeparatorRow(line)) {
            expanded.push(line);
            continue;
        }

        const cells = splitPipeCells(line);
        const splitCells = cells.map(splitCellBreaks);
        const rowCount = splitCells.reduce((max, parts) => Math.max(max, parts.length), 1);
        for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
            expanded.push(markdownRow(splitCells.map(parts => parts[rowIndex] || '')));
        }
    }

    return expanded.join('\n');
}

function normalizeInlineLabeledBlockTable(cells) {
    const groups = [];
    let current = [];

    for (const cell of cells) {
        if (isTableRangeConnectorText(cell)) {
            if (current.length > 0) {
                groups.push(current);
                current = [];
            }
            continue;
        }
        current.push(cell);
    }
    if (current.length > 0) {
        groups.push(current);
    }

    if (groups.length < 2 || groups.some(group => group.length < 2)) {
        return null;
    }

    const columns = groups.map(group => ({
        label: stripTrailingColon(group[0]),
        details: group.slice(1)
    }));
    if (columns.some(column => !column.label || /[.!?]$/.test(column.label))) {
        return null;
    }

    return markdownTableForLabeledColumns(columns);
}

function normalizeInlinePipeTable(text) {
    const cells = splitPipeCells(text);
    if (cells.length < 2) {
        return null;
    }

    const labeledBlockTable = normalizeInlineLabeledBlockTable(cells);
    if (labeledBlockTable) {
        return labeledBlockTable;
    }

    const columnCount = cells.length % 2 === 0 ? 2 : Math.min(3, cells.length);
    if (columnCount < 2 || cells.length < columnCount) {
        return null;
    }

    const rows = [];
    for (let i = 0; i < cells.length; i += columnCount) {
        const row = cells.slice(i, i + columnCount);
        while (row.length < columnCount) row.push('');
        rows.push(row);
    }

    return [
        markdownRow(rows[0]),
        markdownSeparator(columnCount),
        ...rows.slice(1).map(markdownRow)
    ].join('\n');
}

export function normalizeTableReplacementText(text) {
    const normalized = normalizeReplacementText(text).trim();
    if (!normalized) {
        return normalized;
    }
    if (isMarkdownTableText(normalized)) {
        return expandMarkdownTableHtmlBreaks(normalized);
    }

    const lines = normalized
        .split(/\r?\n/)
        .map(line => line.trim())
        .filter(Boolean);

    if (lines.length === 1 && lines[0].includes('|')) {
        return normalizeInlinePipeTable(lines[0]) || normalized;
    }

    const pipeLines = lines.filter(line => line.includes('|'));
    if (pipeLines.length >= 2) {
        const rows = pipeLines.map(splitPipeCells).filter(row => row.length > 0);
        const columnCount = rows.reduce((max, row) => Math.max(max, row.length), 0);
        if (columnCount >= 2) {
            const normalizedRows = rows.map(row => {
                const copy = row.slice(0, columnCount);
                while (copy.length < columnCount) copy.push('');
                return copy;
            });
            return [
                markdownRow(normalizedRows[0]),
                markdownSeparator(columnCount),
                ...normalizedRows.slice(1).map(markdownRow)
            ].join('\n');
        }
    }

    return normalized;
}

function replacementLooksTableLike(text) {
    const normalized = normalizeReplacementText(text).trim();
    return isMarkdownTableText(normalized)
        || (normalized.includes('|') && splitPipeCells(normalized).length >= 2);
}

export function isDiagnosticReplacementText(text) {
    const normalized = normalizeReplacementText(text)
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
    if (!normalized) return false;

    const diagnosticPatterns = [
        /schema requires/,
        /schema.*paragraphindex/,
        /replace_range operation/,
        /replace_paragraph operation/,
        /do not include ["']?originaltext/,
        /do not include .*replacementtext/,
        /originaltext.*replacementtext/,
        /provide only .*paragraphindex.*operation.*content/,
        /none required for replace_range/,
        /invalid replace_range/
    ];

    return diagnosticPatterns.some(pattern => pattern.test(normalized));
}

function isLikelyTableSourceText(text) {
    const normalized = String(text || '').trim();
    if (!normalized) return false;
    if (isTableRangeConnectorText(normalized)) return true;
    if (isLikelyStructuredTableSourceParagraph(normalized)) return true;

    const logicalLines = normalized
        .split(/[\r\n\v]+/)
        .map(line => line.trim())
        .filter(Boolean);

    return logicalLines.length > 1
        && logicalLines.some(line => /:\s*$/.test(line))
        && logicalLines.some(line => /^\[.*\]$/.test(line) || /^\(.*\)$/.test(line));
}

function getLogicalParagraphLines(text) {
    return String(text || '')
        .split(/[\r\n\v]+/)
        .map(line => line.trim())
        .filter(Boolean);
}

function isTableRangeConnectorText(text) {
    return /^(?:and|or|between|with|plus|to)$/i.test(String(text || '').trim());
}

function parseLabeledBlockParagraph(text) {
    const lines = getLogicalParagraphLines(text);
    if (lines.length < 2) return null;

    const firstLine = lines[0];
    const label = stripTrailingColon(firstLine);
    const hasExplicitLabel = /:\s*$/.test(firstLine);
    const looksLikeCompactHeading = label.length > 0
        && label.length <= 80
        && !/[.!?]$/.test(label)
        && !/\s(?:is|are|was|were|shall|must|may|will|can|could|should|would)\s/i.test(label);

    if (!hasExplicitLabel && !looksLikeCompactHeading) return null;

    return {
        label,
        details: lines.slice(1)
    };
}

export function synthesizeMarkdownTableFromSourceRange(paragraphItems, startIndex, endIndex) {
    const items = Array.isArray(paragraphItems) ? paragraphItems : [];
    if (
        !Number.isInteger(startIndex)
        || !Number.isInteger(endIndex)
        || startIndex < 0
        || endIndex < startIndex
        || endIndex >= items.length
    ) {
        return null;
    }

    const columns = [];
    for (let index = startIndex; index <= endIndex; index += 1) {
        const text = String(items[index]?.text || '').trim();
        if (!text || isTableRangeConnectorText(text)) continue;

        const parsed = parseLabeledBlockParagraph(text);
        if (parsed) {
            columns.push(parsed);
        }
    }

    if (columns.length >= 2) {
        return markdownTableForLabeledColumns(columns);
    }

    return null;
}

export function inferTableConversionEndIndex(paragraphItems, startIndex, maxScan = 10) {
    const items = Array.isArray(paragraphItems) ? paragraphItems : [];
    if (!Number.isInteger(startIndex) || startIndex < 0 || startIndex >= items.length) {
        return startIndex;
    }

    let endIndex = startIndex;
    const limit = Math.min(items.length - 1, startIndex + Math.max(1, maxScan));
    for (let index = startIndex + 1; index <= limit; index += 1) {
        const text = String(items[index]?.text || '').trim();
        if (!text) {
            if (endIndex > startIndex) break;
            continue;
        }
        if (!isLikelyTableSourceText(text)) {
            break;
        }
        endIndex = index;
    }

    return endIndex;
}

export function findNearbyParagraphIndexForModifyText(paragraphItems, startIndex, change) {
    const items = Array.isArray(paragraphItems) ? paragraphItems : [];
    if (!Number.isInteger(startIndex) || startIndex < 0 || startIndex >= items.length) {
        return startIndex;
    }

    const originalText = normalizeNeedleText(change?.originalText);
    if (!originalText) {
        return startIndex;
    }

    const textAt = index => String(items[index]?.text ?? '');
    const startText = textAt(startIndex);
    if (startText.trim().length > 0) {
        return startIndex;
    }

    const originalLower = originalText.toLowerCase();
    const maxDistance = 12;
    let bestExact = null;
    let bestInsensitive = null;

    for (let distance = 1; distance <= maxDistance; distance += 1) {
        const candidates = [startIndex - distance, startIndex + distance];
        for (const candidateIndex of candidates) {
            if (candidateIndex < 0 || candidateIndex >= items.length) continue;

            const candidateText = textAt(candidateIndex);
            if (!candidateText.trim()) continue;

            if (candidateText.includes(originalText)) {
                if (bestExact == null) {
                    bestExact = candidateIndex;
                }
                continue;
            }

            if (candidateText.toLowerCase().includes(originalLower)) {
                if (bestInsensitive == null) {
                    bestInsensitive = candidateIndex;
                }
            }
        }

        if (bestExact != null) return bestExact;
    }

    if (bestInsensitive != null) return bestInsensitive;
    return startIndex;
}

function parseParagraphIndex(value) {
    const parsed = Number.parseInt(String(value ?? ''), 10);
    return Number.isInteger(parsed) ? parsed - 1 : null;
}

export async function applyRedlineChangesToWordContext(context, aiChanges, options = {}) {
    const changes = Array.isArray(aiChanges) ? aiChanges : [];
    const logPrefix = options.logPrefix || 'Redline/Shared';
    const onInfo = typeof options.onInfo === 'function'
        ? options.onInfo
        : message => console.log(`[${logPrefix}] ${message}`);
    const onWarn = typeof options.onWarn === 'function'
        ? options.onWarn
        : message => console.warn(`[${logPrefix}] ${message}`);
    const requestedContentKind = options.requestedContentKind || null;
    const requiresTableContent = requestedContentKind === 'table';

    let changesApplied = 0;
    const skipped = [];
    const recordSkip = (change, operation, reason) => {
        skipped.push({
            paragraphIndex: change?.paragraphIndex,
            operation: operation || String(change?.operation || '').trim().toLowerCase(),
            reason
        });
    };

    for (const change of changes) {
        try {
            let operationName = String(change?.operation || '').trim().toLowerCase();
            const startIndex = parseParagraphIndex(change?.paragraphIndex);
            if (!Number.isInteger(startIndex) || startIndex < 0) {
                onWarn(`Invalid start paragraph index: ${change?.paragraphIndex}`);
                continue;
            }

            if (requiresTableContent && operationName === 'modify_text') {
                onWarn(
                    `Skipping modify_text at P${change?.paragraphIndex}: table conversion requests require replace_paragraph or replace_range with markdown table content.`
                );
                continue;
            }

            const paragraphs = context.document.body.paragraphs;
            paragraphs.load('items/text');
            await context.sync();

            const paragraphCount = paragraphs.items.length;

            // Append-at-end: targeting the paragraph one past the last one with a
            // content-bearing operation inserts brand-new paragraph(s) after the
            // document's last paragraph (the only way to add content when there is
            // no trailing blank paragraph to reuse).
            const isAppendAtEnd = startIndex === paragraphCount
                && (operationName === 'replace_paragraph' || operationName === 'replace_range' || operationName === 'edit_paragraph');

            if (startIndex > paragraphCount || (startIndex === paragraphCount && !isAppendAtEnd)) {
                onWarn(`Out-of-range target P${change?.paragraphIndex} (count=${paragraphCount}); no-op.`);
                recordSkip(change, operationName, `paragraph P${change?.paragraphIndex} is out of range (document has ${paragraphCount} paragraphs)`);
                continue;
            }

            if (isAppendAtEnd) {
                const appendContent = normalizeReplacementText(getReplacementText(change, operationName));
                if (!appendContent || !appendContent.trim()) {
                    onWarn(`Append at P${change?.paragraphIndex}: no content to append; no-op.`);
                    recordSkip(change, operationName, 'no content was provided to append at the end of the document');
                    continue;
                }
                if (paragraphCount === 0) {
                    onWarn('Append requested but document has no paragraphs; no-op.');
                    recordSkip(change, operationName, 'document has no paragraphs to append after');
                    continue;
                }
                // Append new paragraphs after the last one using native (tracked) APIs.
                const lastParagraph = paragraphs.items[paragraphCount - 1];
                await insertContentAsNativeParagraphs(context, lastParagraph, appendContent);
                changesApplied += 1;
                onInfo(`Appended new content after P${paragraphCount} (end of document).`);
                continue;
            }

            let effectiveStartIndex = startIndex;
            if (operationName === 'modify_text') {
                effectiveStartIndex = findNearbyParagraphIndexForModifyText(paragraphs.items, startIndex, change);
                if (effectiveStartIndex !== startIndex) {
                    onWarn(
                        `Rebased modify_text from P${startIndex + 1} to P${effectiveStartIndex + 1} based on originalText match.`
                    );
                }
            }

            const normalizedChange = { ...change, operation: operationName };
            if (operationName === 'edit_paragraph' || operationName === 'replace_paragraph' || operationName === 'replace_range') {
                const replacementText = getReplacementText(normalizedChange, operationName);
                const normalizedTableText = normalizeTableReplacementText(replacementText);
                if (replacementLooksTableLike(replacementText) && normalizedTableText !== normalizeReplacementText(replacementText).trim()) {
                    setReplacementText(normalizedChange, operationName, normalizedTableText);
                    onWarn('Normalized pipe-delimited table replacement into markdown table syntax.');
                }

                if (isMarkdownTableText(normalizedTableText) && operationName === 'edit_paragraph') {
                    operationName = 'replace_paragraph';
                    normalizedChange.operation = operationName;
                    normalizedChange.content = normalizedTableText;
                    delete normalizedChange.newContent;
                    onWarn('Promoted table edit_paragraph to replace_paragraph for OOXML table generation.');
                }
            }

            let endIndex = effectiveStartIndex;
            let insertionBeforeStart = false;
            if (operationName === 'replace_range') {
                let requestedEndIndex = parseParagraphIndex(normalizedChange?.endParagraphIndex);
                if (!Number.isInteger(requestedEndIndex) || requestedEndIndex < -1) {
                    if (requiresTableContent || replacementLooksTableLike(getReplacementText(normalizedChange, operationName))) {
                        requestedEndIndex = inferTableConversionEndIndex(paragraphs.items, effectiveStartIndex);
                        normalizedChange.endParagraphIndex = requestedEndIndex + 1;
                        onWarn(
                            `Inferred missing table replace_range end as P${requestedEndIndex + 1}.`
                        );
                    } else {
                        onWarn(`Invalid replace_range endParagraphIndex: ${normalizedChange?.endParagraphIndex}; no-op.`);
                        continue;
                    }
                }

                if (requestedEndIndex === effectiveStartIndex - 1) {
                    insertionBeforeStart = true;
                    endIndex = effectiveStartIndex;
                    onWarn(
                        `Normalizing replace_range insertion-before-target (P${effectiveStartIndex + 1}..P${requestedEndIndex + 1}) to scoped insertion.`
                    );
                } else {
                    endIndex = requestedEndIndex;
                    if (endIndex < effectiveStartIndex || endIndex >= paragraphCount) {
                        onWarn(`Invalid replace_range endParagraphIndex: ${normalizedChange?.endParagraphIndex}; no-op.`);
                        continue;
                    }
                }

                if (!hasStructuralReplacementField(normalizedChange)) {
                    const replacementCandidate = getReplacementText(normalizedChange, operationName);
                    const shouldAttemptTableRecovery = requiresTableContent
                        || isDiagnosticReplacementText(replacementCandidate);
                    const synthesizedTable = synthesizeMarkdownTableFromSourceRange(
                        paragraphs.items,
                        effectiveStartIndex,
                        endIndex
                    );
                    if (shouldAttemptTableRecovery && synthesizedTable) {
                        normalizedChange.content = synthesizedTable;
                        onWarn('Synthesized missing table content from the source paragraph range.');
                    } else {
                        onWarn(
                            `Skipping malformed replace_range at P${effectiveStartIndex + 1}: missing required content.`
                        );
                        continue;
                    }
                }
            } else if (
                operationName === 'replace_paragraph'
                && (requiresTableContent || replacementLooksTableLike(getReplacementText(normalizedChange, operationName)))
            ) {
                const inferredEndIndex = inferTableConversionEndIndex(paragraphs.items, effectiveStartIndex);
                if (inferredEndIndex > effectiveStartIndex) {
                    operationName = 'replace_range';
                    normalizedChange.operation = operationName;
                    normalizedChange.endParagraphIndex = inferredEndIndex + 1;
                    endIndex = inferredEndIndex;
                    onWarn(
                        `Expanded table replacement scope from P${effectiveStartIndex + 1} to P${inferredEndIndex + 1}.`
                    );
                }
            }

            const replacementAfterRangeResolution = getReplacementText(normalizedChange, operationName);
            if (
                (operationName === 'replace_paragraph' || operationName === 'replace_range')
                && isDiagnosticReplacementText(replacementAfterRangeResolution)
            ) {
                const synthesizedTable = synthesizeMarkdownTableFromSourceRange(
                    paragraphs.items,
                    effectiveStartIndex,
                    endIndex
                );
                if (synthesizedTable) {
                    normalizedChange.content = synthesizedTable;
                    onWarn('Replaced diagnostic model output with synthesized table content from the source range.');
                } else {
                    onWarn(
                        `Skipping malformed ${operationName} at P${effectiveStartIndex + 1}: replacement text contains model diagnostics.`
                    );
                    continue;
                }
            }

            if (
                (operationName === 'replace_paragraph' || operationName === 'replace_range')
                && !hasStructuralReplacementField(normalizedChange)
            ) {
                onWarn(
                    `Skipping malformed ${operationName} at P${effectiveStartIndex + 1}: missing required content.`
                );
                continue;
            }

            if (
                requiresTableContent
                && (operationName === 'edit_paragraph' || operationName === 'replace_paragraph' || operationName === 'replace_range')
            ) {
                const finalReplacementText = getReplacementText(normalizedChange, operationName);
                const normalizedTableText = normalizeTableReplacementText(finalReplacementText);
                if (
                    replacementLooksTableLike(finalReplacementText)
                    && normalizedTableText !== normalizeReplacementText(finalReplacementText).trim()
                ) {
                    setReplacementText(normalizedChange, operationName, normalizedTableText);
                    onWarn('Normalized final table replacement into markdown table syntax.');
                }

                const finalContent = getReplacementText(normalizedChange, operationName);
                if (!isMarkdownTableText(normalizeReplacementText(finalContent).trim())) {
                    const synthesizedTable = synthesizeMarkdownTableFromSourceRange(
                        paragraphs.items,
                        effectiveStartIndex,
                        endIndex
                    );
                    if (synthesizedTable) {
                        operationName = 'replace_range';
                        normalizedChange.operation = operationName;
                        normalizedChange.content = synthesizedTable;
                        normalizedChange.endParagraphIndex = endIndex + 1;
                        delete normalizedChange.newContent;
                        onWarn('Replaced non-table model output with synthesized table content from the source range.');
                    } else {
                        onWarn(
                            `Skipping ${operationName} at P${effectiveStartIndex + 1}: table request did not produce markdown table content.`
                        );
                        continue;
                    }
                }
            }

            const startParagraph = paragraphs.items[effectiveStartIndex];
            const scopeParagraphCount = insertionBeforeStart
                ? 1
                : (endIndex - effectiveStartIndex) + 1;

            const converted = toScopedSharedRedlineOperation(normalizedChange, {
                scopeStartText: startParagraph.text || '',
                scopeParagraphCount,
                insertionBeforeStart
            });
            if (!converted.ok) {
                // Empty target paragraph: the OOXML reconciliation engine cannot diff
                // against empty text (Word rejects its output). Insert the content with
                // native, tracked Word APIs instead of skipping.
                const emptyTarget = /target.*empty|empty.*target/i.test(converted.reason || '');
                const insertContent = emptyTarget
                    ? normalizeReplacementText(getReplacementText(normalizedChange, operationName))
                    : null;
                if (insertContent && insertContent.trim()) {
                    onInfo(`Empty target P${change?.paragraphIndex}: inserting content as a native tracked change.`);
                    await insertContentAsNativeParagraphs(context, startParagraph, insertContent, { fillAnchor: true });
                    changesApplied += 1;
                    continue;
                }
                onWarn(`Skipping change: ${converted.reason}`);
                recordSkip(change, operationName, emptyTarget ? 'empty paragraph and no content to insert' : converted.reason);
                continue;
            }

            const applied = await applyWordOperation(
                context,
                converted.operation,
                scopeParagraphCount === 1
                    ? { paragraph: startParagraph }
                    : { paragraph: startParagraph, endParagraph: paragraphs.items[endIndex] },
                {
                    author: options.author,
                    generateRedlines: options.generateRedlines,
                    disableNativeTracking: options.disableNativeTracking,
                    baseTrackingMode: options.baseTrackingMode ?? null,
                    logPrefix,
                    onInfo,
                    onWarn
                }
            );

            if (applied) {
                changesApplied += 1;
            } else {
                onWarn(`No changes produced for change: ${JSON.stringify(normalizedChange)}`);
                recordSkip(change, operationName, `no changes were produced for P${change?.paragraphIndex} (the new content may match the existing text)`);
            }
        } catch (changeError) {
            onWarn(`Failed to apply change ${JSON.stringify(change)}: ${changeError?.message || changeError}`);
            recordSkip(change, undefined, `apply error: ${changeError?.message || changeError}`);
        }
    }

    onInfo(`Total changes applied: ${changesApplied}`);
    return { changesApplied, skipped };
}
