const fs = require('fs');
const path = require('path');
const assetsDir = 'dist/o365/spfx/commands/project/project-upgrade/assets';
if (!fs.existsSync(assetsDir)) {
  fs.mkdirSync(assetsDir);
}
fs.copyFileSync('src/o365/spfx/commands/project/project-upgrade/assets/tab20x20.png', path.join(assetsDir, 'tab20x20.png'));
fs.copyFileSync('src/o365/spfx/commands/project/project-upgrade/assets/tab96x96.png', path.join(assetsDir, 'tab96x96.png'));