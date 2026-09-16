#!/usr/bin/env node
import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = join(here, '..');

export function loadJson(path) {
  return JSON.parse(readFileSync(path, 'utf8'));
}

export function archiveNamesFrom(archive) {
  if (!archive || !Array.isArray(archive.tools)) {
    throw new Error('archive JSON must have a tools array');
  }
  return new Set(archive.tools.map((tool) => tool.name));
}

export function totalCap(budget) {
  const baseline = budget.surfaces.full.total_tokens;
  const pct = budget.thresholds.total_growth_pct;
  return Math.floor((baseline * (100 + pct)) / 100);
}

function isFiniteNumber(value) {
  return typeof value === 'number' && Number.isFinite(value);
}

export function checkTokenBudget(report, budget, archiveNames) {
  const failures = [];
  const perTool = budget.thresholds.per_tool;
  const newTool = budget.thresholds.new_tool;
  const cap = totalCap(budget);
  const tools = report.tools;

  if (!Array.isArray(tools)) {
    throw new Error('report JSON must have a tools array of {name, tokens}');
  }

  for (const tool of tools) {
    const name = tool.name;
    const tokens = tool.tokens;
    if (!isFiniteNumber(tokens)) {
      failures.push(`${name} is missing a numeric tokens value`);
      continue;
    }
    if (tokens > perTool) {
      failures.push(`${name} is ${tokens} tokens (per-tool ceiling ${perTool})`);
    }
    if (!archiveNames.has(name) && tokens > newTool) {
      failures.push(`${name} is a new tool at ${tokens} tokens (new-tool ceiling ${newTool})`);
    }
  }

  const total = report.total;
  if (!isFiniteNumber(total)) {
    failures.push('report total is missing a numeric value');
  } else if (total > cap) {
    failures.push(
      `full total ${total} exceeds ${cap} (${budget.surfaces.full.total_tokens} + ${budget.thresholds.total_growth_pct}%)`
    );
  }

  return failures;
}

function isCli() {
  const entry = process.argv[1];
  if (!entry) return false;
  return import.meta.url === pathToFileURL(resolve(entry)).href;
}

function main() {
  const reportPath = process.argv[2];
  if (!reportPath) {
    console.error('usage: check-token-budget.js <report.json> [token-budget.json] [archive.json]');
    process.exit(2);
  }

  const budgetPath = process.argv[3] || join(repoRoot, 'token-budget.json');
  const archivePath = process.argv[4] || join(here, 'tools-list-2026-09-14.json');

  const report = loadJson(reportPath);
  const budget = loadJson(budgetPath);
  const archiveNames = archiveNamesFrom(loadJson(archivePath));
  const failures = checkTokenBudget(report, budget, archiveNames);

  if (failures.length > 0) {
    for (const failure of failures) {
      console.error(failure);
    }
    process.exit(1);
  }
}

if (isCli()) {
  main();
}
