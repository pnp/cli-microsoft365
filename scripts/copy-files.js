const fs = require('fs');
const path = require('path');

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

const assetsDir = 'dist/m365/spfx/commands/project/project-upgrade/assets';
mkdirNotExistsSync(assetsDir);
fs.copyFileSync('src/m365/spfx/commands/project/project-upgrade/assets/tab20x20.png', path.join(assetsDir, 'tab20x20.png'));
fs.copyFileSync('src/m365/spfx/commands/project/project-upgrade/assets/tab96x96.png', path.join(assetsDir, 'tab96x96.png'));

const paPcfInitAssetsSourceDir = 'src/m365/pa/commands/pcf/pcf-init/assets';
const paPcfInitCmdDir = 'dist/m365/pa/commands/pcf/pcf-init';
const paPcfInitAssetsDir = 'dist/m365/pa/commands/pcf/pcf-init/assets';
mkdirNotExistsSync(paPcfInitCmdDir);
mkdirNotExistsSync(paPcfInitAssetsDir);
getFilePaths(paPcfInitAssetsSourceDir).forEach(file => copyFile(file, paPcfInitAssetsSourceDir, paPcfInitAssetsDir));

const paSolutionInitAssetsSourceDir = 'src/m365/pa/commands/solution/solution-init/assets';
const paSolutionInitCmdDir = 'dist/m365/pa/commands/solution/solution-init';
const paSolutionInitAssetsDir = 'dist/m365/pa/commands/solution/solution-init/assets';
mkdirNotExistsSync(paSolutionInitCmdDir);
mkdirNotExistsSync(paSolutionInitAssetsDir);
getFilePaths(paSolutionInitAssetsSourceDir).forEach(file => copyFile(file, paSolutionInitAssetsSourceDir, paSolutionInitAssetsDir));