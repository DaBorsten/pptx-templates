#!/usr/bin/env node
/**
 * `bun run validate:pptx`: builds the Tier-2 validator image and runs it on the
 * templates and test outputs.
 *
 * A Node wrapper instead of an inline shell command so it runs the same on
 * Linux, macOS, PowerShell and cmd: no `$PWD`, no shell quoting. Docker is
 * called without a shell, and the repo is mounted with `--mount`, which also
 * copes with spaces and Windows drive letters in the path.
 */
import { spawnSync } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const repo = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  '..',
  '..',
);
const image = 'pptx-validator';

const docker = (args) => {
  const result = spawnSync('docker', args, { stdio: 'inherit' });
  if (result.error) {
    console.error(
      result.error.code === 'ENOENT'
        ? 'validate-pptx: docker not found (install Docker and make sure it is on PATH)'
        : `validate-pptx: ${result.error.message}`,
    );
    process.exit(1);
  }
  if (result.status !== 0) process.exit(result.status ?? 1);
};

docker(['build', '-q', '-t', image, path.join(repo, 'tools', 'validate-pptx')]);
docker([
  'run',
  '--rm',
  '--mount',
  `type=bind,source=${repo},target=/repo,readonly`,
  image,
  '--allowlist',
  '/repo/tools/validate-pptx/allowlist.json',
  '/repo/__tests__/pptx-templates',
  '/repo/__tests__/pptx-output',
]);
