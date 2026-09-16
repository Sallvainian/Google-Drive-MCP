// tests/token-cost.test.js
import assert from 'node:assert';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { describe, it } from 'node:test';
import { fileURLToPath } from 'node:url';

const repoRoot = join(dirname(fileURLToPath(import.meta.url)), '..');
const checkerPath = join(repoRoot, 'scripts', 'check-token-budget.js');
const budgetPath = join(repoRoot, 'token-budget.json');
const archivePath = join(repoRoot, 'scripts', 'tools-list-2026-09-14.json');
const measurerPath = join(repoRoot, 'scripts', 'measure-token-cost.py');

function archiveNames() {
  return JSON.parse(readFileSync(archivePath, 'utf8')).tools.map((tool) => tool.name);
}

function reportFromNames(names, tokenForName) {
  const tools = names.map((name, index) => ({ name, tokens: tokenForName(name, index) }));
  return {
    count: tools.length,
    total: tools.reduce((sum, tool) => sum + tool.tokens, 0),
    average: 0,
    tools,
  };
}

function runChecker(report) {
  const dir = mkdtempSync(join(tmpdir(), 'token-cost-'));
  const reportPath = join(dir, 'report.json');
  writeFileSync(reportPath, `${JSON.stringify(report)}\n`);
  const result = spawnSync(process.execPath, [checkerPath, reportPath], {
    encoding: 'utf8',
    cwd: repoRoot,
  });
  rmSync(dir, { recursive: true, force: true });
  return result;
}

describe('token-budget.json', () => {
  it('publishes surfaces.full as 108 / 18200 / 169 with thresholds 1000 / 500 / 10', () => {
    const budget = JSON.parse(readFileSync(budgetPath, 'utf8'));
    assert.deepStrictEqual(Object.keys(budget.surfaces), ['full']);
    assert.deepStrictEqual(budget.surfaces.full, {
      tools: 108,
      total_tokens: 18200,
      average: 169,
    });
    assert.deepStrictEqual(budget.thresholds, {
      per_tool: 1000,
      new_tool: 500,
      total_growth_pct: 10,
    });
  });
});

describe('check-token-budget.js', () => {
  it('exits 0 for archive names under budget including editTableCell at 748', () => {
    const names = archiveNames();
    assert.equal(names.length, 108);
    assert.equal(names[0], 'readGoogleDoc');
    assert.equal(names[names.length - 1], 'send_draft');
    const result = runChecker(
      reportFromNames(names, (name) => {
        if (name === 'editTableCell') return 748;
        if (name === 'applyParagraphStyle') return 562;
        return 1;
      })
    );
    assert.equal(result.status, 0, result.stderr);
  });

  it('exits non-zero when any tool is 1001 tokens', () => {
    const names = archiveNames();
    const result = runChecker(
      reportFromNames(names, (name) => (name === 'editTableCell' ? 1001 : 1))
    );
    assert.notEqual(result.status, 0);
    assert.match(result.stderr, /editTableCell is 1001 tokens \(per-tool ceiling 1000\)/);
  });

  it('exits non-zero when a name not in the archive is 501 tokens', () => {
    const names = archiveNames();
    const report = reportFromNames(names, () => 1);
    report.tools.push({ name: 'brandNewTool', tokens: 501 });
    report.count += 1;
    report.total += 501;
    const result = runChecker(report);
    assert.notEqual(result.status, 0);
    assert.match(result.stderr, /brandNewTool is a new tool at 501 tokens \(new-tool ceiling 500\)/);
  });

  it('exits 0 when a name not in the archive is 500 tokens and total is under cap', () => {
    const names = archiveNames();
    const report = reportFromNames(names, () => 1);
    report.tools.push({ name: 'brandNewTool', tokens: 500 });
    report.count += 1;
    report.total += 500;
    const result = runChecker(report);
    assert.equal(result.status, 0, result.stderr);
  });

  it('exits non-zero when full total is 20021', () => {
    const names = archiveNames();
    const report = reportFromNames(names, (_name, index) => {
      if (index < 20) return 1000;
      if (index === 20) return 21;
      return 0;
    });
    assert.equal(report.total, 20021);
    const result = runChecker(report);
    assert.notEqual(result.status, 0);
    assert.match(result.stderr, /full total 20021 exceeds 20020 \(18200 \+ 10%\)/);
  });

  it('exits 0 when full total is 20020', () => {
    const names = archiveNames();
    const result = runChecker(
      reportFromNames(names, (name, index) => {
        void name;
        if (index < 20) return 1000;
        if (index === 20) return 20;
        return 0;
      })
    );
    assert.equal(result.status, 0, result.stderr);
  });

  it('exits non-zero when a tool has no numeric tokens', () => {
    const names = archiveNames();
    const report = reportFromNames(names, () => 1);
    delete report.tools[0].tokens;
    const result = runChecker(report);
    assert.notEqual(result.status, 0);
    assert.match(result.stderr, /readGoogleDoc is missing a numeric tokens value/);
  });

  it('exits non-zero when report total is not numeric', () => {
    const names = archiveNames();
    const report = reportFromNames(names, () => 1);
    delete report.total;
    const result = runChecker(report);
    assert.notEqual(result.status, 0);
    assert.match(result.stderr, /report total is missing a numeric value/);
  });

  it('does not treat archive names editTableCell 748 or applyParagraphStyle 562 as new tools', () => {
    const names = archiveNames();
    const result = runChecker(
      reportFromNames(names, (name) => {
        if (name === 'editTableCell') return 748;
        if (name === 'applyParagraphStyle') return 562;
        return 1;
      })
    );
    assert.equal(result.status, 0, result.stderr);
    assert.equal(result.stderr.includes('new tool'), false);
  });
});

describe('measure-token-cost.py method lock', () => {
  it('keeps cl100k_base, compact JSON separators, and the three reproduce print lines', () => {
    const source = readFileSync(measurerPath, 'utf8');
    assert.equal(source.includes('tiktoken.get_encoding("cl100k_base")'), true);
    assert.equal(source.includes('json.dumps(t, separators=(",", ":"))'), true);
    assert.equal(source.includes('print(f"TOOLS: {n}")'), true);
    assert.equal(
      source.includes('print(f"TOTAL TOKENS (cl100k_base, compact JSON): {total:,}")'),
      true
    );
    assert.equal(source.includes('print(f"AVERAGE: {total/n:.0f}")'), true);
    assert.equal(source.includes('"total": total'), true);
    assert.equal(source.includes('"tokens": tot'), true);
  });
});
