import fs from 'fs';
import path from 'path';

const mkdirNotExistsSync = (path) => {
  if (!fs.existsSync(path)) {
    fs.mkdirSync(path);
  }
}

const getFilePaths = (folderPath) => {
  const entryPaths = fs.readdirSync(folderPath).map(entry => path.join(folderPath, entry));
  const filePaths = entryPaths.filter(entryPath => fs.statSync(entryPath).isFile());
  const dirPaths = entryPaths.filter(entryPath => !filePaths.includes(entryPath));
  const dirFiles = dirPaths.reduce((prev, curr) => prev.concat(getFilePaths(curr)), []);
  return [...filePaths, ...dirFiles];
};

const copyFile = (file, sourceDir, destinationDir) => {
  const fileName = path.basename(file);
  const filePath = path.relative(sourceDir, path.dirname(file));
  const destinationFilePath = path.join(destinationDir, filePath);

  mkdirNotExistsSync(destinationFilePath);

  fs.copyFileSync(file, path.join(destinationFilePath, fileName));
};

fs.copyFileSync('src/api.d.ts', 'dist/api.d.ts');

const assetsDir = 'dist/m365/spfx/commands/project/project-upgrade/assets';
mkdirNotExistsSync(assetsDir);
fs.copyFileSync('src/m365/spfx/commands/project/project-upgrade/assets/tab20x20.png', path.join(assetsDir, 'tab20x20.png'));
fs.copyFileSync('src/m365/spfx/commands/project/project-upgrade/assets/tab96x96.png', path.join(assetsDir, 'tab96x96.png'));

const spfxPackageGenerateAssetsSourceDir = 'src/m365/spfx/commands/package/package-generate/assets';
const spfxPackageGenerateCmdDir = 'dist/m365/spfx/commands/package/package-generate';
const spfxPackageGenerateAssetsDir = 'dist/m365/spfx/commands/package/package-generate/assets';
mkdirNotExistsSync(spfxPackageGenerateCmdDir);
mkdirNotExistsSync(spfxPackageGenerateAssetsDir);
getFilePaths(spfxPackageGenerateAssetsSourceDir).forEach(file => copyFile(file, spfxPackageGenerateAssetsSourceDir, spfxPackageGenerateAssetsDir));
