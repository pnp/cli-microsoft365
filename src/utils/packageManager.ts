const packageCommands = {
  npm: {
    install: 'npm i -SE',
    installDev: 'npm i -DE',
    uninstall: 'npm un -S',
    uninstallDev: 'npm un -D'
  },
  pnpm: {
    install: 'pnpm i -E',
    installDev: 'pnpm i -DE',
    uninstall: 'pnpm un',
    uninstallDev: 'pnpm un'
  },
  yarn: {
    install: 'yarn add -E',
    installDev: 'yarn add -DE',
    uninstall: 'yarn remove',
    uninstallDev: 'yarn remove'
  }
};

export const packageManager = {
  getPackageManagerCommand(command: string, packageManager: string): string {
    return (packageCommands as any)[packageManager][command];
  },

  mapPackageManagerCommand({ command, packagesDepExact, packagesDevExact, packagesDepUn, packagesDevUn, packageMgr }: {
    command: string, packagesDevExact: string[],
    packagesDepExact: string[], packagesDepUn: string[], packagesDevUn: string[], packageMgr: string
  }): void {
    // matches must be in this particular order to avoid false matches, eg.
    // uninstallDev contains install
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

  reducePackageManagerCommand({ packagesDepExact, packagesDevExact, packagesDepUn, packagesDevUn, packageMgr }: {
    packagesDepExact: string[], packagesDevExact: string[],
    packagesDepUn: string[], packagesDevUn: string[], packageMgr: string
  }): string[] {
    const commandsToExecute: string[] = [];

    // uninstall commands must come first otherwise there is a chance that
    // whatever we recommended to install, will be immediately uninstalled
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

    return commandsToExecute;
  }
};