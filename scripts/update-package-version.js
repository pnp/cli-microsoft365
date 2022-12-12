const fs = require('fs');
const path = require('path');
const packageJson = require('../package.json');

let commitHash = process.argv[process.argv.length-1].substring(0, 7);
// if the commit hash starts with a 0 and consists of only numbers
// prepend an 'a' to match the semver spec and not include a leading 0
// in the build identifier. https://semver.org/#spec-item-9
if (commitHash[0] === '0' && !isNaN(commitHash)) {
  commitHash = `a${commitHash}`;
}
packageJson.version += `-beta.${commitHash}`;
console.log(packageJson.version);
fs.writeFileSync(path.join(path.resolve('.'), 'package.json'), JSON.stringify(packageJson, null, 2));