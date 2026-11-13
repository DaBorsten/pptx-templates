#!/usr/bin/env node
/**
 * `npx pptx-templates <command>` (the package's `bin`). Fork extension.
 */
import { CliIO, runLintCommand, supportsColor } from '../template-lint/cli';
import { runUsageCommand } from '../template-lint/usage-cli';

const USAGE = `Usage: pptx-templates <command> [options]

Commands:
  lint <file.pptx>...    Check the {{…}} tags of templates
  usage <file.pptx>...   Show which data fields the tags read, and which not

Run pptx-templates <command> --help for the options of a command.`;

export const runCli = async (args: string[], io: CliIO): Promise<number> => {
  const [command, ...rest] = args;
  if (command === 'lint') return runLintCommand(rest, io);
  if (command === 'usage') return runUsageCommand(rest, io);
  if (command === '-h' || command === '--help') {
    io.stdout(`${USAGE}\n`);
    return 0;
  }
  if (!command) {
    io.stderr(`No command given.\n\n${USAGE}\n`);
    return 2;
  }
  io.stderr(`Unknown command "${command}".\n\n${USAGE}\n`);
  return 2;
};

if (require.main === module) {
  runCli(process.argv.slice(2), {
    stdout: (text) => process.stdout.write(text),
    stderr: (text) => process.stderr.write(text),
    color: supportsColor(),
    clearScreen: !!process.stdout.isTTY,
  }).then(
    (code) => (process.exitCode = code),
    (error) => {
      process.stderr.write(`${error?.stack ?? error}\n`);
      process.exitCode = 2;
    },
  );
}
