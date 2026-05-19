const packageCommands = {
  npm: {
    install: 'npm i -SE',
    installDev: 'npm i -DE',
    installLockFile: 'npm i',
    uninstall: 'npm un -S',
    uninstallDev: 'npm un -D',
    override: 'npm pkg set',
    removeOverride: 'npm pkg delete'
  },
  pnpm: {
    install: 'pnpm i -E',
    installDev: 'pnpm i -DE',
    installLockFile: 'pnpm i',
    uninstall: 'pnpm un',
    uninstallDev: 'pnpm un',
    override: 'pnpm pkg set',
    removeOverride: 'pnpm pkg delete'
  },
  yarn: {
    install: 'yarn add -E',
    installDev: 'yarn add -DE',
    uninstall: 'yarn remove',
    uninstallDev: 'yarn remove'
    // Yarn is not supported for project upgrade since their CLI does not support setting overrides.
  }
};

export const packageManager = {
  getPackageManagerCommand(command: string, packageManager: string): string {
    return (packageCommands as any)[packageManager][command];
  },

  mapPackageManagerCommand({ command, packagesDepExact, packagesDevExact, packagesDepUn, packagesDevUn, packagesOverride, packagesOverrideRemove, packageMgr }: {
    command: string, packagesDevExact: string[],
    packagesDepExact: string[], packagesDepUn: string[], packagesDevUn: string[], packagesOverride: string[], packagesOverrideRemove: string[], packageMgr: string
  }): void {
    // matches must be in this particular order to avoid false matches, eg.
    // uninstallDev contains install, removeOverride contains override
    if (command.startsWith(`${packageManager.getPackageManagerCommand('removeOverride', packageMgr)} `)) {
      packagesOverrideRemove.push(command.replace(packageManager.getPackageManagerCommand('removeOverride', packageMgr), '').trim());
      return;
    }
    if (command.startsWith(`${packageManager.getPackageManagerCommand('override', packageMgr)} `)) {
      packagesOverride.push(command.replace(packageManager.getPackageManagerCommand('override', packageMgr), '').trim());
      return;
    }
    if (command.startsWith(`${packageManager.getPackageManagerCommand('uninstallDev', packageMgr)} `)) {
      packagesDevUn.push(command.replace(packageManager.getPackageManagerCommand('uninstallDev', packageMgr), '').trim());
      return;
    }
    if (command.startsWith(`${packageManager.getPackageManagerCommand('installDev', packageMgr)} `)) {
      packagesDevExact.push(command.replace(packageManager.getPackageManagerCommand('installDev', packageMgr), '').trim());
      return;
    }
    if (command.startsWith(`${packageManager.getPackageManagerCommand('uninstall', packageMgr)} `)) {
      packagesDepUn.push(command.replace(packageManager.getPackageManagerCommand('uninstall', packageMgr), '').trim());
      return;
    }
    if (command.startsWith(`${packageManager.getPackageManagerCommand('install', packageMgr)} `)) {
      packagesDepExact.push(command.replace(packageManager.getPackageManagerCommand('install', packageMgr), '').trim());
    }
  },

  reducePackageManagerCommand({ packagesDepExact, packagesDevExact, packagesDepUn, packagesDevUn, packagesOverride, packagesOverrideRemove, packageMgr }: {
    packagesDepExact: string[], packagesDevExact: string[],
    packagesDepUn: string[], packagesDevUn: string[], packagesOverride: string[], packagesOverrideRemove: string[], packageMgr: string
  }): string[] {
    const commandsToExecute: string[] = [];

    // removeOverride comes first to clear stale overrides before any install/uninstall
    // uninstall commands must come before install commands otherwise there is a
    // chance that whatever we recommended to install will be immediately uninstalled
    // override (add) comes last so it is applied after all
    // install/uninstall operations have completed
    if (packagesOverrideRemove.length > 0) {
      commandsToExecute.push(`${packageManager.getPackageManagerCommand('removeOverride', packageMgr)} ${packagesOverrideRemove.join(' ')}`);
      // removeOverride only updates package.json; run a plain install to update the lock file
      // only needed when no other install/uninstall commands will already update the lock file
      if (packagesDepUn.length === 0 && packagesDevUn.length === 0 && packagesDepExact.length === 0 && packagesDevExact.length === 0 && packagesOverride.length === 0) {
        commandsToExecute.push(packageManager.getPackageManagerCommand('installLockFile', packageMgr));
      }
    }

    if (packagesDepUn.length > 0) {
      commandsToExecute.push(`${packageManager.getPackageManagerCommand('uninstall', packageMgr)} ${packagesDepUn.join(' ')}`);
    }

    if (packagesDevUn.length > 0) {
      commandsToExecute.push(`${packageManager.getPackageManagerCommand('uninstallDev', packageMgr)} ${packagesDevUn.join(' ')}`);
    }

    if (packagesDepExact.length > 0) {
      commandsToExecute.push(`${packageManager.getPackageManagerCommand('install', packageMgr)} ${packagesDepExact.join(' ')}`);
    }

    if (packagesDevExact.length > 0) {
      commandsToExecute.push(`${packageManager.getPackageManagerCommand('installDev', packageMgr)} ${packagesDevExact.join(' ')}`);
    }

    if (packagesOverride.length > 0) {
      commandsToExecute.push(`${packageManager.getPackageManagerCommand('override', packageMgr)} ${packagesOverride.join(' ')}`);
      // override only updates package.json; run a plain install to update the lock file
      commandsToExecute.push(packageManager.getPackageManagerCommand('installLockFile', packageMgr));
    }

    return commandsToExecute;
  }
};