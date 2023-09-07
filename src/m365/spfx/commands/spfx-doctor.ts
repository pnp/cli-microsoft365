import child_process from 'child_process';
import { satisfies } from 'semver';
import GlobalOptions from '../../../GlobalOptions.js';
import { Logger } from '../../../cli/Logger.js';
import { CheckStatus, formatting } from '../../../utils/formatting.js';
import commands from '../commands.js';
import { BaseProjectCommand } from './project/base-project-command.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  env?: string;
  spfxVersion?: string;
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
  gulpCli: VersionCheck;
  node: VersionCheck;
  sp: SharePointVersion;
  yo: VersionCheck;
}

class SpfxDoctorCommand extends BaseProjectCommand {
  private readonly versions: { [version: string]: SpfxVersionPrerequisites } = {
    '1.0.0': {
      gulpCli: {
        range: '^1 || ^2',
        fix: 'npm i -g gulp-cli@2'
      },
      node: {
        range: '^6',
        fix: 'Install Node.js v6'
      },
      sp: SharePointVersion.All,
      yo: {
        range: '^3',
        fix: 'npm i -g yo@3'
      }
    },
    '1.1.0': {
      gulpCli: {
        range: '^1 || ^2',
        fix: 'npm i -g gulp-cli@2'
      },
      node: {
        range: '^6',
        fix: 'Install Node.js v6'
      },
      sp: SharePointVersion.All,
      yo: {
        range: '^3',
        fix: 'npm i -g yo@3'
      }
    },
    '1.2.0': {
      gulpCli: {
        range: '^1 || ^2',
        fix: 'npm i -g gulp-cli@2'
      },
      node: {
        range: '^6',
        fix: 'Install Node.js v6'
      },
      sp: SharePointVersion.SP2019 | SharePointVersion.SPO,
      yo: {
        range: '^3',
        fix: 'npm i -g yo@3'
      }
    },
    '1.4.0': {
      gulpCli: {
        range: '^1 || ^2',
        fix: 'npm i -g gulp-cli@2'
      },
      node: {
        range: '^6',
        fix: 'Install Node.js v6'
      },
      sp: SharePointVersion.SP2019 | SharePointVersion.SPO,
      yo: {
        range: '^3',
        fix: 'npm i -g yo@3'
      }
    },
    '1.4.1': {
      gulpCli: {
        range: '^1 || ^2',
        fix: 'npm i -g gulp-cli@2'
      },
      node: {
        range: '^6 || ^8',
        fix: 'Install Node.js v8'
      },
      sp: SharePointVersion.SP2019 | SharePointVersion.SPO,
      yo: {
        range: '^3',
        fix: 'npm i -g yo@3'
      }
    },
    '1.5.0': {
      gulpCli: {
        range: '^1 || ^2',
        fix: 'npm i -g gulp-cli@2'
      },
      node: {
        range: '^6 || ^8',
        fix: 'Install Node.js v8'
      },
      sp: SharePointVersion.SPO,
      yo: {
        range: '^3',
        fix: 'npm i -g yo@3'
      }
    },
    '1.5.1': {
      gulpCli: {
        range: '^1 || ^2',
        fix: 'npm i -g gulp-cli@2'
      },
      node: {
        range: '^6 || ^8',
        fix: 'Install Node.js v8'
      },
      sp: SharePointVersion.SPO,
      yo: {
        range: '^3',
        fix: 'npm i -g yo@3'
      }
    },
    '1.6.0': {
      gulpCli: {
        range: '^1 || ^2',
        fix: 'npm i -g gulp-cli@2'
      },
      node: {
        range: '^6 || ^8',
        fix: 'Install Node.js v8'
      },
      sp: SharePointVersion.SPO,
      yo: {
        range: '^3',
        fix: 'npm i -g yo@3'
      }
    },
    '1.7.0': {
      gulpCli: {
        range: '^1 || ^2',
        fix: 'npm i -g gulp-cli@2'
      },
      node: {
        range: '^8',
        fix: 'Install Node.js v8'
      },
      sp: SharePointVersion.SPO,
      yo: {
        range: '^3',
        fix: 'npm i -g yo@3'
      }
    },
    '1.7.1': {
      gulpCli: {
        range: '^1 || ^2',
        fix: 'npm i -g gulp-cli@2'
      },
      node: {
        range: '^8',
        fix: 'Install Node.js v8'
      },
      sp: SharePointVersion.SPO,
      yo: {
        range: '^3',
        fix: 'npm i -g yo@3'
      }
    },
    '1.8.0': {
      gulpCli: {
        range: '^1 || ^2',
        fix: 'npm i -g gulp-cli@2'
      },
      node: {
        range: '^8',
        fix: 'Install Node.js v8'
      },
      sp: SharePointVersion.SPO,
      yo: {
        range: '^3',
        fix: 'npm i -g yo@3'
      }
    },
    '1.8.1': {
      gulpCli: {
        range: '^1 || ^2',
        fix: 'npm i -g gulp-cli@2'
      },
      node: {
        range: '^8',
        fix: 'Install Node.js v8'
      },
      sp: SharePointVersion.SPO,
      yo: {
        range: '^3',
        fix: 'npm i -g yo@3'
      }
    },
    '1.8.2': {
      gulpCli: {
        range: '^1 || ^2',
        fix: 'npm i -g gulp-cli@2'
      },
      node: {
        range: '^8 || ^10',
        fix: 'Install Node.js v10'
      },
      sp: SharePointVersion.SPO,
      yo: {
        range: '^3',
        fix: 'npm i -g yo@3'
      }
    },
    '1.9.0': {
      gulpCli: {
        range: '^1 || ^2',
        fix: 'npm i -g gulp-cli@2'
      },
      node: {
        range: '^8 || ^10',
        fix: 'Install Node.js v10'
      },
      sp: SharePointVersion.SPO,
      yo: {
        range: '^3',
        fix: 'npm i -g yo@3'
      }
    },
    '1.9.1': {
      gulpCli: {
        range: '^1 || ^2',
        fix: 'npm i -g gulp-cli@2'
      },
      node: {
        range: '^10',
        fix: 'Install Node.js v10'
      },
      sp: SharePointVersion.SPO,
      yo: {
        range: '^3',
        fix: 'npm i -g yo@3'
      }
    },
    '1.10.0': {
      gulpCli: {
        range: '^1 || ^2',
        fix: 'npm i -g gulp-cli@2'
      },
      node: {
        range: '^10',
        fix: 'Install Node.js v10'
      },
      sp: SharePointVersion.SPO,
      yo: {
        range: '^3',
        fix: 'npm i -g yo@3'
      }
    },
    '1.11.0': {
      gulpCli: {
        range: '^1 || ^2',
        fix: 'npm i -g gulp-cli@2'
      },
      node: {
        range: '^10',
        fix: 'Install Node.js v10'
      },
      sp: SharePointVersion.SPO,
      yo: {
        range: '^3',
        fix: 'npm i -g yo@3'
      }
    },
    '1.12.0': {
      gulpCli: {
        range: '^1 || ^2',
        fix: 'npm i -g gulp-cli@2'
      },
      node: {
        range: '^12',
        fix: 'Install Node.js v12'
      },
      sp: SharePointVersion.SPO,
      yo: {
        range: '^3',
        fix: 'npm i -g yo@3'
      }
    },
    '1.12.1': {
      gulpCli: {
        range: '^1 || ^2',
        fix: 'npm i -g gulp-cli@2'
      },
      node: {
        range: '^12 || ^14',
        fix: 'Install Node.js v12 or v14'
      },
      sp: SharePointVersion.SPO,
      yo: {
        range: '^3',
        fix: 'npm i -g yo@3'
      }
    },
    '1.13.0': {
      gulpCli: {
        range: '^1 || ^2',
        fix: 'npm i -g gulp-cli@2'
      },
      node: {
        range: '^12 || ^14',
        fix: 'Install Node.js v12 or v14'
      },
      sp: SharePointVersion.SPO,
      yo: {
        range: '^4',
        fix: 'npm i -g yo@4'
      }
    },
    '1.13.1': {
      gulpCli: {
        range: '^1 || ^2',
        fix: 'npm i -g gulp-cli@2'
      },
      node: {
        range: '^12 || ^14',
        fix: 'Install Node.js v12 or v14'
      },
      sp: SharePointVersion.SPO,
      yo: {
        range: '^4',
        fix: 'npm i -g yo@4'
      }
    },
    '1.14.0': {
      gulpCli: {
        range: '^1 || ^2',
        fix: 'npm i -g gulp-cli@2'
      },
      node: {
        range: '^12 || ^14',
        fix: 'Install Node.js v12 or v14'
      },
      sp: SharePointVersion.SPO,
      yo: {
        range: '^4',
        fix: 'npm i -g yo@4'
      }
    },
    '1.15.0': {
      gulpCli: {
        range: '^1 || ^2',
        fix: 'npm i -g gulp-cli@2'
      },
      node: {
        range: '^12.13 || ^14.15 || ^16.13',
        fix: 'Install Node.js v12.13, v14.15, v16.13 or higher'
      },
      sp: SharePointVersion.SPO,
      yo: {
        range: '^4',
        fix: 'npm i -g yo@4'
      }
    },
    '1.15.2': {
      gulpCli: {
        range: '^1 || ^2',
        fix: 'npm i -g gulp-cli@2'
      },
      node: {
        range: '^12.13 || ^14.15 || ^16.13',
        fix: 'Install Node.js v12.13, v14.15, v16.13 or higher'
      },
      sp: SharePointVersion.SPO,
      yo: {
        range: '^4',
        fix: 'npm i -g yo@4'
      }
    },
    '1.16.0': {
      gulpCli: {
        range: '^1 || ^2',
        fix: 'npm i -g gulp-cli@2'
      },
      node: {
        range: '>=16.13.0 <17.0.0',
        fix: 'Install Node.js >=16.13.0 <17.0.0'
      },
      sp: SharePointVersion.SPO,
      yo: {
        range: '^4',
        fix: 'npm i -g yo@4'
      }
    },
    '1.16.1': {
      gulpCli: {
        range: '^1 || ^2',
        fix: 'npm i -g gulp-cli@2'
      },
      node: {
        range: '>=16.13.0 <17.0.0',
        fix: 'Install Node.js >=16.13.0 <17.0.0'
      },
      sp: SharePointVersion.SPO,
      yo: {
        range: '^4',
        fix: 'npm i -g yo@4'
      }
    },
    '1.17.0': {
      gulpCli: {
        range: '^1 || ^2',
        fix: 'npm i -g gulp-cli@2'
      },
      node: {
        range: '>=16.13.0 <17.0.0',
        fix: 'Install Node.js >=16.13.0 <17.0.0'
      },
      sp: SharePointVersion.SPO,
      yo: {
        range: '^4',
        fix: 'npm i -g yo@4'
      }
    },
    '1.17.1': {
      gulpCli: {
        range: '^1 || ^2',
        fix: 'npm i -g gulp-cli@2'
      },
      node: {
        range: '>=16.13.0 <17.0.0',
        fix: 'Install Node.js >=16.13.0 <17.0.0'
      },
      sp: SharePointVersion.SPO,
      yo: {
        range: '^4',
        fix: 'npm i -g yo@4'
      }
    },
    '1.17.2': {
      gulpCli: {
        range: '^1 || ^2',
        fix: 'npm i -g gulp-cli@2'
      },
      node: {
        range: '>=16.13.0 <17.0.0',
        fix: 'Install Node.js >=16.13.0 <17.0.0'
      },
      sp: SharePointVersion.SPO,
      yo: {
        range: '^4',
        fix: 'npm i -g yo@4'
      }
    },
    '1.17.3': {
      gulpCli: {
        range: '^1 || ^2',
        fix: 'npm i -g gulp-cli@2'
      },
      node: {
        range: '>=16.13.0 <17.0.0',
        fix: 'Install Node.js >=16.13.0 <17.0.0'
      },
      sp: SharePointVersion.SPO,
      yo: {
        range: '^4',
        fix: 'npm i -g yo@4'
      }
    },
    '1.17.4': {
      gulpCli: {
        range: '^1 || ^2',
        fix: 'npm i -g gulp-cli@2'
      },
      node: {
        range: '>=16.13.0 <17.0.0',
        fix: 'Install Node.js >=16.13.0 <17.0.0'
      },
      sp: SharePointVersion.SPO,
      yo: {
        range: '^4',
        fix: 'npm i -g yo@4'
      }
    },
    '1.18.0-rc.1': {
      gulpCli: {
        range: '^1 || ^2',
        fix: 'npm i -g gulp-cli@2'
      },
      node: {
        range: '>=16.13.0 <17.0.0 || >=18.17.1 <19.0.0',
        fix: 'Install Node.js >=16.13.0 <17.0.0 || >=18.17.1 <19.0.0'
      },
      sp: SharePointVersion.SPO,
      yo: {
        range: '^4',
        fix: 'npm i -g yo@4'
      }
    }
  };

  protected get allowedOutputs(): string[] {
    return ['text'];
  }

  public get name(): string {
    return commands.DOCTOR;
  }

  public get description(): string {
    return 'Verifies environment configuration for using the specific version of the SharePoint Framework';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        env: args.options.env,
        spfxVersion: args.options.spfxVersion
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-e, --env [env]',
        autocomplete: ['sp2016', 'sp2019', 'spo']
      },
      {
        option: '-v, --spfxVersion [spfxVersion]',
        autocomplete: Object.keys(this.versions)
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.env) {
          const sp: SharePointVersion | undefined = this.spVersionStringToEnum(args.options.env);
          if (!sp) {
            return `${args.options.env} is not a valid SharePoint version. Valid versions are sp2016, sp2019 or spo`;
          }
        }

        if (args.options.spfxVersion) {
          if (!this.versions[args.options.spfxVersion]) {
            return `${args.options.spfxVersion} is not a supported SharePoint Framework version. Supported versions are ${Object.keys(this.versions).join(', ')}`;
          }
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (!args.options.output) {
      args.options.output = 'text';
    }

    this.projectRootPath = this.getProjectRoot(process.cwd());

    await logger.log(' ');
    await logger.log('CLI for Microsoft 365 SharePoint Framework doctor');
    await logger.log('Verifying configuration of your system for working with the SharePoint Framework');
    await logger.log(' ');

    let spfxVersion: string = '';
    let prerequisites: SpfxVersionPrerequisites;
    const fixes: string[] = [];

    try {
      spfxVersion = args.options.spfxVersion ?? await this.getSharePointFrameworkVersion(logger);

      if (!spfxVersion) {
        await logger.log(formatting.getStatus(CheckStatus.Failure, `SharePoint Framework`));
        throw `SharePoint Framework not found`;
      }

      prerequisites = this.versions[spfxVersion];
      if (!prerequisites) {
        await logger.log(formatting.getStatus(CheckStatus.Failure, `SharePoint Framework v${spfxVersion}`));
        throw `spfx doctor doesn't support SPFx v${spfxVersion} at this moment`;
      }

      if (args.options.spfxVersion) {
        await this.checkSharePointFrameworkVersion(args.options.spfxVersion, fixes, logger);
      }
      else {
        // spfx was detected and if we are here, it means that we support it
        await logger.log(formatting.getStatus(CheckStatus.Success, `SharePoint Framework v${spfxVersion}`));
      }

      await this.checkSharePointCompatibility(spfxVersion, prerequisites, args, fixes, logger);
      await this.checkNodeVersion(prerequisites, fixes, logger);
      await this.checkYo(prerequisites, fixes, logger);
      await this.checkGulp(fixes, logger);
      await this.checkGulpCli(prerequisites, fixes, logger);
      await this.checkTypeScript(fixes, logger);

      if (fixes.length > 0) {
        await logger.log(' ');
        await logger.log('Recommended fixes:');
        await logger.log(' ');
        for (const f of fixes) {
          await logger.log(`- ${f}`);
        }
        await logger.log(' ');
      }
    }
    catch (err: any) {
      await logger.log(' ');

      if (fixes.length > 0) {
        await logger.log('Recommended fixes:');
        await logger.log(' ');
        for (const f of fixes) {
          await logger.log(`- ${f}`);
        }
        await logger.log(' ');
      }
      this.handleRejectedPromise(err);
    }
  }

  private async checkSharePointCompatibility(spfxVersion: string, prerequisites: SpfxVersionPrerequisites, args: CommandArgs, fixes: string[], logger: Logger): Promise<void> {
    if (args.options.env) {
      const sp: SharePointVersion = this.spVersionStringToEnum(args.options.env) as SharePointVersion;
      if ((prerequisites.sp & sp) === sp) {
        await logger.log(formatting.getStatus(CheckStatus.Success, `Supported in ${SharePointVersion[sp]}`));
        return;
      }

      await logger.log(formatting.getStatus(CheckStatus.Failure, `Not supported in ${SharePointVersion[sp]}`));
      fixes.push(`Use SharePoint Framework v${(sp === SharePointVersion.SP2016 ? '1.1' : '1.4.1')}`);
      throw `SharePoint Framework v${spfxVersion} is not supported in ${SharePointVersion[sp]}`;
    }
  }

  private async checkNodeVersion(prerequisites: SpfxVersionPrerequisites, fixes: string[], logger: Logger): Promise<void> {
    const nodeVersion: string = this.getNodeVersion();
    this.checkStatus('Node', nodeVersion, prerequisites.node, OptionalOrRequired.Required, fixes, logger);
  }

  private async checkSharePointFrameworkVersion(spfxVersionRequested: string, fixes: string[], logger: Logger): Promise<void> {
    let spfxVersionDetected = await this.getSPFxVersionFromYoRcFile(logger);
    if (!spfxVersionDetected) {
      spfxVersionDetected = await this.getPackageVersion('@microsoft/generator-sharepoint', PackageSearchMode.GlobalOnly, HandlePromise.Continue, logger);
    }
    const versionCheck: VersionCheck = {
      range: spfxVersionRequested,
      fix: `npm i -g @microsoft/generator-sharepoint@${spfxVersionRequested}`
    };
    if (spfxVersionDetected) {
      this.checkStatus(`SharePoint Framework`, spfxVersionDetected, versionCheck, OptionalOrRequired.Required, fixes, logger);
    }
    else {
      await logger.log(formatting.getStatus(CheckStatus.Failure, `SharePoint Framework v${spfxVersionRequested} not found`));
      fixes.push(versionCheck.fix);
    }
  }

  private async checkYo(prerequisites: SpfxVersionPrerequisites, fixes: string[], logger: Logger): Promise<void> {
    const yoVersion: string = await this.getPackageVersion('yo', PackageSearchMode.GlobalOnly, HandlePromise.Continue, logger);
    if (yoVersion) {
      this.checkStatus('yo', yoVersion, prerequisites.yo, OptionalOrRequired.Required, fixes, logger);
    }
    else {
      await logger.log(formatting.getStatus(CheckStatus.Failure, `yo not found`));
      fixes.push(prerequisites.yo.fix);
    }
  }

  private async checkGulpCli(prerequisites: SpfxVersionPrerequisites, fixes: string[], logger: Logger): Promise<void> {
    const gulpCliVersion: string = await this.getPackageVersion('gulp-cli', PackageSearchMode.GlobalOnly, HandlePromise.Continue, logger);
    if (gulpCliVersion) {
      this.checkStatus('gulp-cli', gulpCliVersion, prerequisites.gulpCli, OptionalOrRequired.Required, fixes, logger);
    }
    else {
      await logger.log(formatting.getStatus(CheckStatus.Failure, `gulp-cli not found`));
      fixes.push(prerequisites.gulpCli.fix);
    }
  }

  private async checkGulp(fixes: string[], logger: Logger): Promise<void> {
    const gulpVersion: string = await this.getPackageVersion('gulp', PackageSearchMode.GlobalOnly, HandlePromise.Continue, logger);
    if (gulpVersion) {
      await logger.log(formatting.getStatus(CheckStatus.Failure, `gulp should be removed`));
      fixes.push('npm un -g gulp');
    }
  }

  private async checkTypeScript(fixes: string[], logger: Logger): Promise<void> {
    const typeScriptVersion: string = await this.getPackageVersion('typescript', PackageSearchMode.LocalOnly, HandlePromise.Continue, logger);
    if (typeScriptVersion) {
      await logger.log(formatting.getStatus(CheckStatus.Failure, `typescript v${typeScriptVersion} installed in the project`));
      fixes.push('npm un typescript');
    }
    else {
      await logger.log(formatting.getStatus(CheckStatus.Success, `bundled typescript used`));
    }
  }

  private spVersionStringToEnum(sp: string): SharePointVersion | undefined {
    return (<any>SharePointVersion)[sp.toUpperCase()];
  }

  private async getSPFxVersionFromYoRcFile(logger: Logger): Promise<string | undefined> {
    if (this.projectRootPath !== null) {
      const spfxVersion = this.getProjectVersion();
      if (spfxVersion) {
        if (this.debug) {
          await logger.logToStderr(`SPFx version retrieved from .yo-rc.json file. Retrieved version: ${spfxVersion}`);
        }
        return spfxVersion;
      }
    }
    return undefined;
  }

  private async getSharePointFrameworkVersion(logger: Logger): Promise<string> {
    let spfxVersion = await this.getSPFxVersionFromYoRcFile(logger);
    if (spfxVersion) {
      return spfxVersion;
    }
    try {
      spfxVersion = await this.getPackageVersion('@microsoft/sp-core-library', PackageSearchMode.LocalOnly, HandlePromise.Fail, logger);
      if (this.debug) {
        await logger.logToStderr(`Found @microsoft/sp-core-library@${spfxVersion}`);
      }
      return spfxVersion;
    }
    catch {
      if (this.debug) {
        await logger.logToStderr(`@microsoft/sp-core-library not found. Search for @microsoft/generator-sharepoint local or global...`);
      }

      try {
        return await this.getPackageVersion('@microsoft/generator-sharepoint', PackageSearchMode.LocalAndGlobal, HandlePromise.Fail, logger);
      }
      catch (error: any) {
        if (this.debug) {
          await logger.logToStderr('@microsoft/generator-sharepoint not found');
        }

        if (error && error.indexOf('ENOENT') > -1) {
          throw 'npm not found';
        }
        else {
          return '';
        }
      }
    }
  }

  private async getPackageVersion(packageName: string, searchMode: PackageSearchMode, handlePromise: HandlePromise, logger: Logger): Promise<string> {
    const args: string[] = ['ls', packageName, '--depth=0', '--json'];
    if (searchMode === PackageSearchMode.GlobalOnly) {
      args.push('-g');
    }

    let version: string;
    try {
      version = await this.getPackageVersionFromNpm(args, logger);
    }
    catch {
      if (searchMode === PackageSearchMode.LocalAndGlobal) {
        args.push('-g');
        version = await this.getPackageVersionFromNpm(args, logger);
      }
      else {
        version = '';
      }
    }

    if (version) {
      return version;
    }
    else {
      if (handlePromise === HandlePromise.Continue) {
        return '';
      }
      else {
        throw new Error();
      }
    }
  }

  private getPackageVersionFromNpm(args: string[], logger: Logger): Promise<string> {
    return new Promise<string>(async (resolve: (version: string) => void, reject: (error: string) => void): Promise<void> => {
      const packageName: string = args[1];

      if (this.debug) {
        await logger.logToStderr(`Executing npm: ${args.join(' ')}...`);
      }

      child_process.exec(`npm ${args.join(' ')}`, (err: child_process.ExecException | null, stdout: string): void => {
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
        catch (ex: any) {
          return reject(ex);
        }
      });
    });
  }

  private getNodeVersion(): string {
    return process.version.substr(1);
  }

  private async checkStatus(what: string, versionFound: string, versionCheck: VersionCheck, optionalOrRequired: OptionalOrRequired, fixes: string[], logger: Logger): Promise<void> {
    if (versionFound) {
      if (satisfies(versionFound, versionCheck.range)) {
        await logger.log(formatting.getStatus(CheckStatus.Success, `${what} v${versionFound}`));
      }
      else {
        await logger.log(formatting.getStatus(CheckStatus.Failure, `${what} v${versionFound} found, v${versionCheck.range} required`));
        fixes.push(versionCheck.fix);
      }
    }
  }
}

export default new SpfxDoctorCommand();
