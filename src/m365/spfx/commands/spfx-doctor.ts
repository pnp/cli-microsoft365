import commands from '../commands';
import GlobalOptions from '../../../GlobalOptions';
import { CommandError, CommandOption, CommandValidate, CommandTypes } from '../../../Command';
import * as child_process from 'child_process';
import AnonymousCommand from '../../base/AnonymousCommand';
import { satisfies } from 'semver';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  env?: string;
}

/**
 * Has the particular check passed or failed
 */
enum CheckStatus {
  Success,
  Failure
}

/**
 * Where to search for the particular npm package: only in the current project,
 * in global packages or both
 */
enum PackageSearchMode {
  LocalOnly,
  GlobalOnly,
  LocalAndGlobal
}

/**
 * Is the particular check optional or required
 */
enum OptionalOrRequired {
  Optional,
  Required
}

/**
 * Should the method continue or fail on a rejected Promise
 */
enum HandlePromise {
  Fail,
  Continue
}

interface VersionCheck {
  /**
   * Required version range in semver
   */
  range: string;
  /**
   * What to do to fix it if the required range isn't met
   */
  fix: string;
}

/**
 * Versions of SharePoint that support SharePoint Framework
 */
enum SharePointVersion {
  SP2016 = 1 << 0,
  SP2019 = 1 << 1,
  SPO = 1 << 2,
  All = ~(~0 << 3)
}

interface SpfxVersionPrerequisites {
  node: VersionCheck;
  npm: VersionCheck;
  react: VersionCheck;
  sp: SharePointVersion
}

class SpfxDoctorCommand extends AnonymousCommand {
  private readonly versions: {
    [version: string]: SpfxVersionPrerequisites
  } = {
      '1.0.0': {
        node: {
          range: '^6.0.0',
          fix: 'Install Node.js v6'
        },
        npm: {
          range: '^3.0.0',
          fix: 'npm i -g npm@3'
        },
        react: {
          range: '^15',
          fix: 'npm i react@15'
        },
        sp: SharePointVersion.All
      },
      '1.1.0': {
        node: {
          range: '^6.0.0',
          fix: 'Install Node.js v6'
        },
        npm: {
          range: '^3.0.0 || ^4.0.0',
          fix: 'npm i -g npm@4'
        },
        react: {
          range: '^15',
          fix: 'npm i react@15'
        },
        sp: SharePointVersion.All
      },
      '1.2.0': {
        node: {
          range: '^6.0.0',
          fix: 'Install Node.js v6'
        },
        npm: {
          range: '^3.0.0 || ^4.0.0',
          fix: 'npm i -g npm@4'
        },
        react: {
          range: '^15',
          fix: 'npm i react@15'
        },
        sp: SharePointVersion.SP2019 | SharePointVersion.SPO
      },
      '1.4.0': {
        node: {
          range: '^6.0.0',
          fix: 'Install Node.js v6'
        },
        npm: {
          range: '^3.0.0 || ^4.0.0',
          fix: 'npm i -g npm@4'
        },
        react: {
          range: '^15',
          fix: 'npm i react@15'
        },
        sp: SharePointVersion.SP2019 | SharePointVersion.SPO
      },
      '1.4.1': {
        node: {
          range: '^6.0.0 || ^8.0.0',
          fix: 'Install Node.js v8'
        },
        npm: {
          range: '^3.0.0 || ^4.0.0',
          fix: 'npm i -g npm@4'
        },
        react: {
          range: '^15',
          fix: 'npm i react@15'
        },
        sp: SharePointVersion.SP2019 | SharePointVersion.SPO
      },
      '1.5.0': {
        node: {
          range: '^6.0.0 || ^8.0.0',
          fix: 'Install Node.js v8'
        },
        npm: {
          range: '^3.0.0',
          fix: 'npm i -g npm@3'
        },
        react: {
          range: '^15',
          fix: 'npm i react@15'
        },
        sp: SharePointVersion.SPO
      },
      '1.5.1': {
        node: {
          range: '^6.0.0 || ^8.0.0',
          fix: 'Install Node.js v8'
        },
        npm: {
          range: '^5.0.0',
          fix: 'npm i -g npm@5'
        },
        react: {
          range: '^15',
          fix: 'npm i react@15'
        },
        sp: SharePointVersion.SPO
      },
      '1.6.0': {
        node: {
          range: '^6.0.0 || ^8.0.0',
          fix: 'Install Node.js v8'
        },
        npm: {
          range: '^5.0.0',
          fix: 'npm i -g npm@5'
        },
        react: {
          range: '^15',
          fix: 'npm i react@15'
        },
        sp: SharePointVersion.SPO
      },
      '1.7.0': {
        node: {
          range: '^8.0.0',
          fix: 'Install Node.js v8'
        },
        npm: {
          range: '^5.0.0 || ^6.0.0',
          fix: 'npm i -g npm@6'
        },
        react: {
          range: '16.3.2',
          fix: 'npm i react@16.3.2'
        },
        sp: SharePointVersion.SPO
      },
      '1.7.1': {
        node: {
          range: '^8.0.0',
          fix: 'Install Node.js v8'
        },
        npm: {
          range: '^5.0.0 || ^6.0.0',
          fix: 'npm i -g npm@6'
        },
        react: {
          range: '16.3.2',
          fix: 'npm i react@16.3.2'
        },
        sp: SharePointVersion.SPO
      },
      '1.8.0': {
        node: {
          range: '^8.0.0',
          fix: 'Install Node.js v8'
        },
        npm: {
          range: '^5.0.0 || ^6.0.0',
          fix: 'npm i -g npm@6'
        },
        react: {
          range: '16.7.0',
          fix: 'npm i react@16.7.0'
        },
        sp: SharePointVersion.SPO
      },
      '1.8.1': {
        node: {
          range: '^8.0.0',
          fix: 'Install Node.js v8'
        },
        npm: {
          range: '^5.0.0 || ^6.0.0',
          fix: 'npm i -g npm@6'
        },
        react: {
          range: '16.7.0',
          fix: 'npm i react@16.7.0'
        },
        sp: SharePointVersion.SPO
      },
      '1.8.2': {
        node: {
          range: '^8.0.0 || ^10.0.0',
          fix: 'Install Node.js v10'
        },
        npm: {
          range: '^5.0.0 || ^6.0.0',
          fix: 'npm i -g npm@6'
        },
        react: {
          range: '16.7.0',
          fix: 'npm i react@16.7.0'
        },
        sp: SharePointVersion.SPO
      },
      '1.9.0': {
        node: {
          range: '^8.0.0 || ^10.0.0',
          fix: 'Install Node.js v10'
        },
        npm: {
          range: '^5.0.0 || ^6.0.0',
          fix: 'npm i -g npm@6'
        },
        react: {
          range: '16.8.5',
          fix: 'npm i react@16.8.5'
        },
        sp: SharePointVersion.SPO
      },
      '1.9.1': {
        node: {
          range: '^10.0.0',
          fix: 'Install Node.js v10'
        },
        npm: {
          range: '^5.0.0 || ^6.0.0',
          fix: 'npm i -g npm@6'
        },
        react: {
          range: '16.8.5',
          fix: 'npm i react@16.8.5'
        },
        sp: SharePointVersion.SPO
      },
      '1.10.0': {
        node: {
          range: '^10.0.0',
          fix: 'Install Node.js v10'
        },
        npm: {
          range: '^5.0.0 || ^6.0.0',
          fix: 'npm i -g npm@6'
        },
        react: {
          range: '16.8.5',
          fix: 'npm i react@16.8.5'
        },
        sp: SharePointVersion.SPO
      },
      '1.11.0': {
        node: {
          range: '^10.0.0',
          fix: 'Install Node.js v10'
        },
        npm: {
          range: '^5.0.0 || ^6.0.0',
          fix: 'npm i -g npm@6'
        },
        react: {
          range: '16.8.5',
          fix: 'npm i react@16.8.5'
        },
        sp: SharePointVersion.SPO
      }
    };

  public get name(): string {
    return commands.DOCTOR;
  }

  public get description(): string {
    return 'Verifies environment configuration for using the specific version of the SharePoint Framework';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    cmd.log(' ');
    cmd.log('CLI for Microsoft 365 SharePoint Framework doctor');
    cmd.log('Verifying configuration of your system for working with the SharePoint Framework');
    cmd.log(' ');

    let spfxVersion: string = '';
    let prerequisites: SpfxVersionPrerequisites;
    const fixes: string[] = [];

    this
      .getSharePointFrameworkVersion(cmd)
      .then((_spfxVersion: string): Promise<void> => {
        if (!_spfxVersion) {
          cmd.log(this.getStatus(CheckStatus.Failure, `SharePoint Framework`));
          return Promise.reject(`SharePoint Framework not found`);
        }

        spfxVersion = _spfxVersion;

        prerequisites = this.versions[spfxVersion];
        if (!prerequisites) {
          cmd.log(this.getStatus(CheckStatus.Failure, `SharePoint Framework v${spfxVersion}`));
          return Promise.reject(`spfx doctor doesn't support SPFx v${spfxVersion} at this moment`);
        }

        cmd.log(this.getStatus(CheckStatus.Success, `SharePoint Framework v${spfxVersion}`));
        return Promise.resolve();
      })
      .then(_ => this.checkSharePointCompatibility(spfxVersion, prerequisites, args, fixes, cmd))
      .then(_ => this.checkNodeVersion(prerequisites, fixes, cmd))
      .then(_ => this.checkNpmVersion(prerequisites, fixes, cmd))
      .then(_ => this.checkYo(fixes, cmd))
      .then(_ => this.checkGulp(fixes, cmd))
      .then(_ => this.checkReact(prerequisites, fixes, cmd))
      .then(_ => this.checkTypeScript(fixes, cmd))
      .then(_ => {
        if (fixes.length > 0) {
          cmd.log(' ');
          cmd.log('Recommended fixes:');
          cmd.log(' ');
          fixes.forEach(f => cmd.log(`- ${f}`));
          cmd.log(' ');
        }

        cb();
      })
      .catch((error: string): void => {
        cmd.log(' ');

        if (fixes.length > 0) {
          cmd.log('Recommended fixes:');
          cmd.log(' ');
          fixes.forEach(f => cmd.log(`- ${f}`));
          cmd.log(' ');
        }

        cb(new CommandError(error));
      });
  }

  private checkSharePointCompatibility(spfxVersion: string, prerequisites: SpfxVersionPrerequisites, args: CommandArgs, fixes: string[], cmd: CommandInstance): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: string) => void): void => {
      if (args.options.env) {
        const sp: SharePointVersion = this.spVersionStringToEnum(args.options.env) as SharePointVersion;
        if ((prerequisites.sp & sp) === sp) {
          cmd.log(this.getStatus(CheckStatus.Success, `Supported in ${SharePointVersion[sp]}`));
          resolve();
        }
        else {
          cmd.log(this.getStatus(CheckStatus.Failure, `Not supported in ${SharePointVersion[sp]}`));
          fixes.push(`Use SharePoint Framework v${(sp === SharePointVersion.SP2016 ? '1.1' : '1.4.1')}`);
          reject(`SharePoint Framework v${spfxVersion} is not supported in ${SharePointVersion[sp]}`);
        }
      }
      else {
        resolve();
      }
    });
  }

  private checkNodeVersion(prerequisites: SpfxVersionPrerequisites, fixes: string[], cmd: CommandInstance): Promise<void> {
    return Promise
      .resolve(this.getNodeVersion())
      .then((nodeVersion: string): void => {
        this.checkStatus('Node', nodeVersion, prerequisites.node, OptionalOrRequired.Required, fixes, cmd);
      });
  }

  private checkNpmVersion(prerequisites: SpfxVersionPrerequisites, fixes: string[], cmd: CommandInstance): Promise<void> {
    return this
      .getNpmVersion()
      .then((npmVersion: string): void => {
        this.checkStatus('npm', npmVersion, prerequisites.npm, OptionalOrRequired.Required, fixes, cmd);
      }, (error: string): Promise<void> => {
        cmd.log(this.getStatus(CheckStatus.Failure, error));
        return Promise.reject(error);
      });
  }

  private checkYo(fixes: string[], cmd: CommandInstance): Promise<void> {
    return this
      .getPackageVersion('yo', PackageSearchMode.GlobalOnly, HandlePromise.Continue, cmd)
      .then((yoVersion: string): void => {
        if (yoVersion) {
          cmd.log(this.getStatus(CheckStatus.Success, `yo v${yoVersion}`));
        }
        else {
          cmd.log(this.getStatus(CheckStatus.Failure, `yo not found`));
          fixes.push('npm i -g yo');
        }
      });
  }

  private checkGulp(fixes: string[], cmd: CommandInstance): Promise<void> {
    return this
      .getPackageVersion('gulp', PackageSearchMode.GlobalOnly, HandlePromise.Continue, cmd)
      .then((gulpVersion: string): void => {
        if (gulpVersion) {
          cmd.log(this.getStatus(CheckStatus.Success, `gulp v${gulpVersion}`));
        }
        else {
          cmd.log(this.getStatus(CheckStatus.Failure, `gulp not found`));
          fixes.push('npm i -g gulp');
        }
      });
  }

  private checkReact(prerequisites: SpfxVersionPrerequisites, fixes: string[], cmd: CommandInstance): Promise<void> {
    return this
      .getPackageVersion('react', PackageSearchMode.LocalOnly, HandlePromise.Continue, cmd)
      .then((reactVersion: string): void => {
        this.checkStatus('react', reactVersion, prerequisites.react, OptionalOrRequired.Optional, fixes, cmd);
      });
  }

  private checkTypeScript(fixes: string[], cmd: CommandInstance): Promise<void> {
    return this
      .getPackageVersion('typescript', PackageSearchMode.LocalOnly, HandlePromise.Continue, cmd)
      .then((typeScriptVersion: string): void => {
        if (typeScriptVersion) {
          cmd.log(this.getStatus(CheckStatus.Failure, `typescript v${typeScriptVersion} installed in the project`));
          fixes.push('npm un typescript');
        }
        else {
          cmd.log(this.getStatus(CheckStatus.Success, `bundled typescript used`));
        }
      });
  }

  private spVersionStringToEnum(sp: string): SharePointVersion | undefined {
    return (<any>SharePointVersion)[sp.toUpperCase()];
  }

  private getSharePointFrameworkVersion(cmd: CommandInstance): Promise<string> {
    return new Promise<string>((resolve: (version: string) => void, reject: (error: string) => void): void => {
      if (this.debug) {
        cmd.log('Detecting SharePoint Framework version based on @microsoft/sp-core-library local...');
      }

      this
        .getPackageVersion('@microsoft/sp-core-library', PackageSearchMode.LocalOnly, HandlePromise.Fail, cmd)
        .then((version: string): Promise<string> => {
          if (this.debug) {
            cmd.log(`Found @microsoft/sp-core-library@${version}`);
          }

          return Promise.resolve(version);
        })
        .catch((): Promise<string> => {
          if (this.debug) {
            cmd.log(`@microsoft/sp-core-library not found. Search for @microsoft/generator-sharepoint local or global...`);
          }

          return this.getPackageVersion('@microsoft/generator-sharepoint', PackageSearchMode.LocalAndGlobal, HandlePromise.Fail, cmd);
        })
        .then((version: string): void => {
          resolve(version);
        })
        .catch((error?: string): void => {
          if (this.debug) {
            cmd.log('@microsoft/generator-sharepoint not found');
          }

          if (error && error.indexOf('ENOENT') > -1) {
            reject('npm not found');
          }
          else {
            resolve('');
          }
        });
    });
  }

  private getPackageVersion(packageName: string, searchMode: PackageSearchMode, handlePromise: HandlePromise, cmd: CommandInstance): Promise<string> {
    return new Promise<string>((resolve: (version: string) => void, reject: (err?: any) => void): void => {
      const args: string[] = ['ls', packageName, '--depth=0', '--json'];
      if (searchMode === PackageSearchMode.GlobalOnly) {
        args.push('-g');
      }

      this
        .getPackageVersionFromNpm(args, cmd)
        .then((version: string): Promise<string> => {
          return Promise.resolve(version);
        })
        .catch((): Promise<string> => {
          if (searchMode === PackageSearchMode.LocalAndGlobal) {
            args.push('-g');
            return this.getPackageVersionFromNpm(args, cmd);
          }
          else {
            return Promise.resolve('');
          }
        })
        .then((version: string): void => {
          if (version) {
            resolve(version);
          }
          else {
            if (handlePromise === HandlePromise.Continue) {
              resolve('');
            }
            else {
              reject();
            }
          }
        })
        .catch((err: string): void => {
          reject(err);
        });
    });
  }

  private getPackageVersionFromNpm(args: string[], cmd: CommandInstance): Promise<string> {
    return new Promise<string>((resolve: (version: string) => void, reject: (error: string) => void): void => {
      const packageName: string = args[1];

      if (this.debug) {
        cmd.log(`Executing npm: ${args.join(' ')}...`);
      }

      child_process.execFile(/^win/.test(process.platform) ? 'npm.cmd' : 'npm', args, (err: child_process.ExecException | null, stdout: string, stderr: string): void => {
        if (err) {
          reject(err.message);
        }

        const responseString: string = stdout;
        try {
          const packageInfo: {
            dependencies?: {
              [packageName: string]: {
                version: string;
              };
            };
          } = JSON.parse(responseString);
          if (packageInfo.dependencies &&
            packageInfo.dependencies[packageName]) {
            resolve(packageInfo.dependencies[packageName].version);
          }
          else {
            reject('Package not found');
          }
        }
        catch (ex) {
          return reject(ex);
        }
      });
    });
  }

  private getNodeVersion(): string {
    return process.version.substr(1);
  }

  private getNpmVersion(): Promise<string> {
    return new Promise<string>((resolve: (version: string) => void, reject: (error: string) => void): void => {
      child_process.execFile(/^win/.test(process.platform) ? 'npm.cmd' : 'npm', ['-v'], (err: child_process.ExecException | null, stdout: string, stderr: string): void => {
        if (err) {
          return reject('npm not found');
        }

        resolve(stdout.trim());
      });
    });
  }

  private checkStatus(what: string, versionFound: string, versionCheck: VersionCheck, optionalOrRequired: OptionalOrRequired, fixes: string[], cmd: CommandInstance): void {
    if (!versionFound) {
      // TODO: we might need this code in the future if SPFx introduces required
      // prerequisites with a specific version
      // if (optionalOrRequired === OptionalOrRequired.Required) {
      //   cmd.log(this.getStatus(CheckStatus.Failure, `${what} not found, v${versionCheck.range} required`));
      //   fixes.push(versionCheck.fix);
      // }
    }
    else {
      if (satisfies(versionFound, versionCheck.range)) {
        cmd.log(this.getStatus(CheckStatus.Success, `${what} v${versionFound}`));
      }
      else {
        cmd.log(this.getStatus(CheckStatus.Failure, `${what} v${versionFound} found, v${versionCheck.range} required`));
        fixes.push(versionCheck.fix);
      }
    }
  }

  private getStatus(result: CheckStatus, message: string) {
    const primarySupported: boolean = process.platform !== 'win32' ||
      process.env.CI === 'true' ||
      process.env.TERM === 'xterm-256color';
    const success: string = primarySupported ? '✔' : '√';
    const failure: string = primarySupported ? '✖' : '×';
    return `${result === CheckStatus.Success ? chalk.green(success) : chalk.red(failure)} ${message}`;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-e, --env [env]',
        description: 'Version of SharePoint for which to check compatibility: sp2016|sp2019|spo',
        autocomplete: ['sp2016', 'sp2019', 'spo']
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public types(): CommandTypes | undefined {
    return {
      string: ['e', 'env']
    };
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (args.options.env) {
        const sp: SharePointVersion | undefined = this.spVersionStringToEnum(args.options.env);
        if (!sp) {
          return `${args.options.env} is not a valid SharePoint version. Valid versions are sp2016, sp2019 or spo`;
        }
      }

      return true;
    };
  }
}

module.exports = new SpfxDoctorCommand();