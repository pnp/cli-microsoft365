#!/usr/bin/env zx
$.verbose = false;

console.log('Deprecate beta versions of the @pnp/cli-microsoft365 npm package on npm');
const version = await question('Version of the package to deprecate: ');
const otp = await question('One-time password: ');
const allVersions = JSON.parse(await $`npm view @pnp/cli-microsoft365 versions -json`);
const versionsToDeprecate = allVersions.filter(v => v !== null && v.startsWith(`${version}-beta`));

if (versionsToDeprecate.length === 0) {
  console.log(`No versions matching ${version}-beta found`);
  process.exit();
}

for (let i = 0; i < versionsToDeprecate.length; i++) {
  const v = versionsToDeprecate[i];
  console.log(`Deprecating ${v}...`);
  try {
    await $`npm deprecate "@pnp/cli-microsoft365@${v}" "Preview version released" --otp=${otp}`;
  }
  catch (err) {
    console.error(chalk.red(err.stderr));
  }
}
