import type { VcsChangeInfo, VcsConfig } from '@docusaurus/types';
import { getVcsPreset } from '@docusaurus/utils';
import { execFile } from 'node:child_process';
import { availableParallelism } from 'node:os';
import { relative, resolve, sep } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const defaultVcs = getVcsPreset('git-ad-hoc');
const siteDir = resolve(__dirname, '..', '..');
const repoRoot = resolve(siteDir, '..');
const versionedDocsPrefix = 'versioned_docs/version-';

// Same default limit as Docusaurus' git-ad-hoc strategy to avoid spawning
// one git process per page at once
const gitConcurrency = availableParallelism() * 4;
let activeGitCommands = 0;
const pendingGitCommands: (() => void)[] = [];

async function runGit(args: string[]): Promise<string> {
  if (activeGitCommands >= gitConcurrency) {
    await new Promise<void>(resolveSlot => pendingGitCommands.push(resolveSlot));
  }
  else {
    activeGitCommands++;
  }

  try {
    const { stdout } = await execFileAsync('git', ['-c', 'log.showSignature=false', ...args], { cwd: repoRoot });
    return stdout;
  }
  finally {
    const next = pendingGitCommands.shift();
    if (next) {
      next();
    }
    else {
      activeGitCommands--;
    }
  }
}

// versioned_docs is generated at build time and is not tracked by git, so
// Docusaurus can't find the last update info. We resolve it from the docs
// in the git tag that the version was created from instead.
async function getVersionedFileInfo(filePath: string, age: 'oldest' | 'newest'): Promise<VcsChangeInfo | null | undefined> {
  const relativePath = relative(siteDir, filePath).split(sep).join('/');
  if (!relativePath.startsWith(versionedDocsPrefix)) {
    return undefined;
  }

  const [tag, ...segments] = relativePath.substring(versionedDocsPrefix.length).split('/');
  const sourcePath = `docs/docs/${segments.join('/')}`;
  const ageArgs = age === 'oldest' ? ['--follow', '--diff-filter=A'] : [];

  try {
    const stdout = await runGit(['log', '--max-count=1', '--format=%ct%n%an', ...ageArgs, tag, '--', sourcePath]);
    const [timestamp, author] = stdout.trim().split('\n');
    return timestamp ? { timestamp: parseInt(timestamp, 10) * 1000, author } : null;
  }
  catch {
    return null;
  }
}

export const versionedDocsVcs: VcsConfig = {
  initialize: params => defaultVcs.initialize(params),
  getFileCreationInfo: async filePath => (await getVersionedFileInfo(filePath, 'oldest')) ?? defaultVcs.getFileCreationInfo(filePath),
  getFileLastUpdateInfo: async filePath => (await getVersionedFileInfo(filePath, 'newest')) ?? defaultVcs.getFileLastUpdateInfo(filePath)
};
