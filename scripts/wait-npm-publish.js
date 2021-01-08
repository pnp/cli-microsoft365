const { exec } = require('child_process');

const args = process.argv.slice(2);
const tag = args[0]
const version = args[1];

const waitForPublish = function (origResolve, origReject, tag, version) {
  return new Promise((resolve, reject) => {
    resolve = (origResolve) ? origResolve : resolve;
    reject = (origReject) ? origReject : reject;
    exec(`npm view @pnp/cli-microsoft365@${tag} version`, (error, stdout, stderr) => {
      if (error) {
        reject(error);
      }
      if (stdout.trim() === version) {
        resolve(version)
      } else {
        setTimeout(() => waitForPublish(resolve, reject, tag, version), 1000)
      }
    })
  })
}

waitForPublish(null, null, tag, version)
  .then(version => console.log(`DONE - ${version}`))
  .catch(error => console.error(`ERROR - ${error}`))
