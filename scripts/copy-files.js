const fs = require('fs');
const path = require('path');
const assetsDir = 'dist/o365/spfx/commands/project/project-upgrade/assets';
if (!fs.existsSync(assetsDir)) {
  fs.mkdirSync(assetsDir);
}
fs.copyFileSync('src/o365/spfx/commands/project/project-upgrade/assets/tab20x20.png', path.join(assetsDir, 'tab20x20.png'));
fs.copyFileSync('src/o365/spfx/commands/project/project-upgrade/assets/tab96x96.png', path.join(assetsDir, 'tab96x96.png'));

const paPcfAssetsSourceDir = 'src/o365/pa/commands/pcf/pcf-init/assets';
const paPcfAssetsDir = 'dist/o365/pa/commands/pcf/pcf-init/assets';
if (!fs.existsSync(paPcfAssetsDir)) {
  fs.mkdirSync(paPcfAssetsDir);
}

const getFilePaths = (folderPath) => {
  const entryPaths = fs.readdirSync(folderPath).map(entry => path.join(folderPath, entry));
  const filePaths = entryPaths.filter(entryPath => fs.statSync(entryPath).isFile());
  const dirPaths = entryPaths.filter(entryPath => !filePaths.includes(entryPath));
  const dirFiles = dirPaths.reduce((prev, curr) => prev.concat(getFilePaths(curr)), []);
  return [...filePaths, ...dirFiles];
};

getFilePaths(paPcfAssetsSourceDir).forEach(file => {
  const fileName = path.basename(file);
  const filePath = path.relative(paPcfAssetsSourceDir, path.dirname(file));
  const destinationFilePath = path.join(paPcfAssetsDir, filePath);

  if (!fs.existsSync(destinationFilePath)) {
    fs.mkdirSync(destinationFilePath);
  }

  fs.copyFileSync(file, path.join(destinationFilePath, fileName));
});