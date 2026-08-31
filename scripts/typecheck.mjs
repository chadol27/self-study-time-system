import { cp, mkdir, readdir, rm } from 'node:fs/promises';
import { spawnSync } from 'node:child_process';
import path from 'node:path';

const root = path.resolve(import.meta.dirname, '..');
const sourceDirectory = path.join(root, 'src');
const temporaryDirectory = path.join(root, '.typecheck');

await rm(temporaryDirectory, { force: true, recursive: true });
await mkdir(temporaryDirectory);

try {
  const files = (await readdir(sourceDirectory)).filter((file) => file.endsWith('.gs'));
  await Promise.all(files.map((file) => cp(
    path.join(sourceDirectory, file),
    path.join(temporaryDirectory, `${path.basename(file, '.gs')}.js`)
  )));

  const tsc = path.join(root, 'node_modules', 'typescript', 'bin', 'tsc');
  const result = spawnSync(process.execPath, [tsc, '--project', 'tsconfig.json'], {
    cwd: root,
    stdio: 'inherit'
  });
  process.exitCode = result.status ?? 1;
} finally {
  await rm(temporaryDirectory, { force: true, recursive: true });
}
