const exec = require('child_process').exec;
const os = require('os');
const path = require('path');

function cb(error, stdout, stderr) {
  if (error) {
    console.error(error);
  }

  if (stdout) {
    console.log(stdout);
  }

  if (stderr) {
    console.error(stderr);
  }
}

const distPath = path.join(process.cwd(), 'dist');

switch (os.type()) {
  case 'Linux':
  case 'Darwin':
    exec(`rm -rf "${distPath}"`, cb);
    break;
  case 'Windows_NT':
    exec(`rd /s /q "${distPath}"`, cb);
    break;
  default:
    console.log(`Unsupported OS ${os.type()}. Please delete the 'dist' folder manually.`);
}