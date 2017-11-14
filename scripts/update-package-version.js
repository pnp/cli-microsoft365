const fs = require('fs');
const path = require('path');
const packageJson = require('../package.json');
packageJson.version += `-beta.${process.argv[process.argv.length-1].substr(0, 7)}`;
console.log(packageJson.version);
fs.writeFileSync(path.join(path.resolve('.'), 'package.json'), JSON.stringify(packageJson, null, 2));