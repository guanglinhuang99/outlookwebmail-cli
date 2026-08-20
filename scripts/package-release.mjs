#!/usr/bin/env node

import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { chmod, copyFile, cp, mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { basename, join, resolve } from 'node:path';
import process from 'node:process';

const projectRoot = resolve(import.meta.dirname, '..');
const packageJson = JSON.parse(await readFile(join(projectRoot, 'package.json'), 'utf8'));
const version = packageJson.version;
const releaseDirectory = join(projectRoot, 'release');
const stagingDirectory = join(releaseDirectory, '.staging');
const packageFilename = `${packageJson.name}-${version}.tgz`;

function run(command, args, options = {}) {
  const result = spawnSync(command, args, {
    cwd: projectRoot,
    encoding: 'utf8',
    stdio: options.capture ? 'pipe' : 'inherit',
    ...options,
  });
  if (result.error) throw result.error;
  if (result.status !== 0) {
    throw new Error(`${command} ${args.join(' ')} failed with exit code ${result.status}.`);
  }
  return result.stdout ?? '';
}

async function createBundle(platform) {
  const bundleName = `webmail-cli-${version}-${platform}`;
  const bundleDirectory = join(stagingDirectory, bundleName);
  await mkdir(bundleDirectory, { recursive: true });
  await copyFile(join(releaseDirectory, packageFilename), join(bundleDirectory, packageFilename));
  await cp(join(projectRoot, 'packaging', platform), bundleDirectory, { recursive: true });
  if (platform === 'macos') await chmod(join(bundleDirectory, 'install.command'), 0o755);

  const archivePath = join(releaseDirectory, `${bundleName}.zip`);
  run('zip', ['-q', '-r', archivePath, bundleName], { cwd: stagingDirectory });
  return archivePath;
}

async function sha256(path) {
  return createHash('sha256').update(await readFile(path)).digest('hex');
}

if (process.platform === 'win32') {
  throw new Error('发布打包脚本需要 macOS、Linux 或 WSL 中的 zip 命令；生成的 Windows ZIP 可在 Windows 上安装。');
}

await rm(releaseDirectory, { recursive: true, force: true });
await mkdir(stagingDirectory, { recursive: true });
run(process.platform === 'win32' ? 'npm.cmd' : 'npm', [
  'pack', '--json', '--pack-destination', releaseDirectory,
], {
  env: { ...process.env, npm_config_cache: join(stagingDirectory, '.npm-cache') },
});

const packagePath = join(releaseDirectory, packageFilename);
const archives = [packagePath, await createBundle('macos'), await createBundle('windows')];
const checksums = [];
for (const path of archives) checksums.push(`${await sha256(path)}  ${basename(path)}`);
await writeFile(join(releaseDirectory, 'SHA256SUMS.txt'), `${checksums.join('\n')}\n`, 'utf8');
await rm(stagingDirectory, { recursive: true, force: true });

process.stdout.write(`Release packages created in ${releaseDirectory}:\n`);
for (const path of archives) process.stdout.write(`- ${basename(path)}\n`);
process.stdout.write('- SHA256SUMS.txt\n');
