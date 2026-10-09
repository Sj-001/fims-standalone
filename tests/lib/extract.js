// App.jsx and server/lib/sheets.js define most of this app's real logic as `const` closures inside one
// big component function (or, on the server, inside module scope) — there's no build step that splits
// them into separately-importable modules. Rather than refactor the app just to make it testable (real
// risk of breaking working code for no behavioral gain), these tests extract the EXACT source text of a
// named function straight out of the real file by anchor matching, then evaluate it with `new Function`.
// This means a test always runs against whatever the file currently says — it cannot silently drift out
// of sync with a refactor the way a hand-copied reimplementation could.
const fs = require('fs');

function readLines(filePath) {
  return fs.readFileSync(filePath, 'utf8').split('\n');
}

// Extracts lines from the first line matching startAnchor through the first SUBSEQUENT line matching
// endAnchor, inclusive of both by default. An anchor is a plain substring (line.includes) or a function
// (line) => boolean for anything a substring can't pin down precisely.
function extractBlock(filePath, startAnchor, endAnchor, opts = {}) {
  const lines = readLines(filePath);
  const matches = (line, anchor) => (typeof anchor === 'function' ? anchor(line) : line.includes(anchor));
  const startIdx = lines.findIndex(l => matches(l, startAnchor));
  if (startIdx === -1) throw new Error(`extractBlock [${filePath}]: start anchor not found: ${startAnchor}`);
  const endIdx = lines.findIndex((l, i) => i > startIdx && matches(l, endAnchor));
  if (endIdx === -1) throw new Error(`extractBlock [${filePath}]: end anchor not found: ${endAnchor}`);
  const endInclusive = opts.endExclusive ? endIdx - 1 : endIdx;
  return lines.slice(startIdx, endInclusive + 1).join('\n');
}

// Builds a callable function from one or more extracted source blocks. `paramNames`/`paramValues` inject
// whatever closure variables the real component would normally provide (registerState, abbreviations,
// etc.) — order must match between the two arrays. Concatenation order of `sourceBlocks` doesn't matter:
// none of these blocks call each other at const-declaration time, only when actually invoked, by which
// point the whole combined body has finished running top to bottom.
function buildFromSource(sourceBlocks, returnName, paramNames = [], paramValues = []) {
  const body = (Array.isArray(sourceBlocks) ? sourceBlocks.join('\n\n') : sourceBlocks) + `\nreturn ${returnName};`;
  const fn = new Function(...paramNames, body);
  return fn(...paramValues);
}

// Returns the single line matching an anchor (same matching rules as extractBlock) — for a one-line
// `const x = ...;` declaration where there's no separate "end" line to anchor on.
function extractLine(filePath, anchor) {
  const lines = readLines(filePath);
  const matches = (line, a) => (typeof a === 'function' ? a(line) : line.includes(a));
  const idx = lines.findIndex(l => matches(l, anchor));
  if (idx === -1) throw new Error(`extractLine [${filePath}]: anchor not found: ${anchor}`);
  return lines[idx];
}

module.exports = { readLines, extractBlock, extractLine, buildFromSource };
