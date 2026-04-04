import { execSync } from 'node:child_process';
import { existsSync, rmSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(__dirname, '..');
const docsRoot = resolve(repoRoot, 'docs');

// Get the last git tag
let lastTag;
try {
  lastTag = execSync('git describe --tags --abbrev=0', {
    encoding: 'utf-8',
    cwd: repoRoot
  }).trim();
}
catch {
  console.log('No git tags found. Skipping stable version preparation.');
  process.exit(0);
}

console.log(`Creating stable version from tag: ${lastTag}`);

// Clean any existing versioned files
for (const p of ['versioned_docs', 'versioned_sidebars', 'versions.json']) {
  const fullPath = resolve(docsRoot, p);
  if (existsSync(fullPath)) {
    rmSync(fullPath, { recursive: true, force: true });
  }
}

try {
  // Temporarily replace docs content with the tagged version
  rmSync(resolve(docsRoot, 'docs'), { recursive: true });
  execSync(`git restore --source="${lastTag}" -- docs/docs/ docs/src/config/sidebars.ts`, {
    cwd: repoRoot
  });

  // Use Docusaurus to create the versioned snapshot
  execSync(`npx docusaurus docs:version "${lastTag}"`, {
    cwd: docsRoot,
    stdio: 'inherit'
  });

  console.log(`Stable version created successfully from tag ${lastTag}`);
}
finally {
  // Restore current branch's docs
  execSync('git restore -- docs/docs/ docs/src/config/sidebars.ts', {
    cwd: repoRoot
  });
  // Clean up any files from the tag that don't exist on current branch
  execSync('git clean -fd docs/docs/', {
    cwd: repoRoot
  });
}
