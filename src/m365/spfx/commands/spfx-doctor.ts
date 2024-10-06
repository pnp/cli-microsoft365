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

export interface SpfxDoctorCheck {
  check: string;
  passed: boolean;
  message: string;
  version?: string;
  fix?: string;
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
    '1.18.0': {
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
    },
    '1.18.1': {
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
    },
    '1.18.2': {
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
        range: '^4 || ^5',
        fix: 'npm i -g yo@5'
      }
    },
    '1.19.0': {
      gulpCli: {
        range: '^1 || ^2 || ^3',
        fix: 'npm i -g gulp-cli@3'
      },
      node: {
        range: '>=18.17.1 <19.0.0',
        fix: 'Install Node.js >=18.17.1 <19.0.0'
      },
      sp: SharePointVersion.SPO,
      yo: {
        range: '^4 || ^5',
        fix: 'npm i -g yo@5'
      }
    },
    '1.20.0': {
      gulpCli: {
        range: '^1 || ^2 || ^3',
        fix: 'npm i -g gulp-cli@3'
      },
      node: {
        range: '>=18.17.1 <19.0.0',
        fix: 'Install Node.js >=18.17.1 <19.0.0'
      },
      sp: SharePointVersion.SPO,
      yo: {
        range: '^4 || ^5',
        fix: 'npm i -g yo@5'
      }
    }
  };

  private output: string = '';
  private resultsObject: SpfxDoctorCheck[] = [];
  private logger!: Logger;

  protected get allowedOutputs(): string[] {
    return ['text', 'json'];
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

    this.output = args.options.output;
    this.projectRootPath = this.getProjectRoot(process.cwd());
    this.logger = logger;

    await this.logMessage(' ');
    await this.logMessage('CLI for Microsoft 365 SharePoint Framework doctor');
    await this.logMessage('Verifying configuration of your system for working with the SharePoint Framework');
    await this.logMessage(' ');

    let spfxVersion: string = '';
    let prerequisites: SpfxVersionPrerequisites;

    try {
      spfxVersion = args.options.spfxVersion ?? await this.getSharePointFrameworkVersion();

      if (!spfxVersion) {
        await this.logMessage(formatting.getStatus(CheckStatus.Failure, `SharePoint Framework`));
        this.resultsObject.push({
          check: 'SharePoint Framework',
          passed: false,
          message: `SharePoint Framework not found`
        });
        throw `SharePoint Framework not found`;
      }

      prerequisites = this.versions[spfxVersion];

      if (!prerequisites) {
        const message = `spfx doctor doesn't support SPFx v${spfxVersion} at this moment`;
        this.resultsObject.push({
          check: 'SharePoint Framework',
          passed: true,
          version: spfxVersion,
          message: message
        });
        await this.logMessage(formatting.getStatus(CheckStatus.Failure, `SharePoint Framework v${spfxVersion}`));
        throw message;
      }
      else {
        this.resultsObject.push({
          check: 'SharePoint Framework',
          passed: true,
          version: spfxVersion,
          message: `SharePoint Framework v${spfxVersion} valid.`
        });
      }

      if (args.options.spfxVersion) {
        await this.checkSharePointFrameworkVersion(args.options.spfxVersion);
      }
      else {
        // spfx was detected and if we are here, it means that we support it
        const message = `SharePoint Framework v${spfxVersion}`;
        this.resultsObject.push({
          check: 'SharePoint Framework',
          passed: true,
          version: spfxVersion,
          message: message
        });
        await this.logMessage(formatting.getStatus(CheckStatus.Success, message));
      }

      await this.checkSharePointCompatibility(spfxVersion, prerequisites, args);
      await this.checkNodeVersion(prerequisites);
      await this.checkYo(prerequisites);
      await this.checkGulp();
      await this.checkGulpCli(prerequisites);
      await this.checkTypeScript();

      if (this.resultsObject.some(y => y.fix !== undefined)) {
        await this.logMessage('Recommended fixes:');
        await this.logMessage(' ');
        for (const f of this.resultsObject.filter(y => y.fix !== undefined)) {
          await this.logMessage(`- ${f.fix}`);
        }
        await this.logMessage(' ');
      }
    }
    catch (err: any) {
      await this.logMessage(' ');

      if (this.resultsObject.some(y => y.fix !== undefined)) {
        await this.logMessage('Recommended fixes:');
        await this.logMessage(' ');
        for (const f of this.resultsObject.filter(y => y.fix !== undefined)) {
          await this.logMessage(`- ${f.fix}`);
        }
        await this.logMessage(' ');
      }

      if (this.output === 'text') {
        this.handleRejectedPromise(err);
      }
    }
    finally {
      if (args.options.output === 'json' && this.resultsObject.length > 0) {
        await logger.log(this.resultsObject);
      }
    }
  }

  private async logMessage(message: string): Promise<void> {
    if (this.output === 'json') {
      await this.logger.logToStderr(message);
    }
    else {
      await this.logger.log(message);
    }
  }

  private async checkSharePointCompatibility(spfxVersion: string, prerequisites: SpfxVersionPrerequisites, args: CommandArgs): Promise<void> {
    if (args.options.env) {
      const sp: SharePointVersion = this.spVersionStringToEnum(args.options.env) as SharePointVersion;
      if ((prerequisites.sp & sp) === sp) {
        const message = `Supported in ${SharePointVersion[sp]}`;
        this.resultsObject.push({
          check: 'env',
          passed: true,
          message: message,
          version: args.options.env
        });
        await this.logMessage(formatting.getStatus(CheckStatus.Success, message));
        return;
      }
      const fix = `Use SharePoint Framework v${(sp === SharePointVersion.SP2016 ? '1.1' : '1.4.1')}`;
      const message = `Not supported in ${SharePointVersion[sp]}`;
      this.resultsObject.push({
        check: 'env',
        passed: false,
        fix: fix,
        message: message,
        version: args.options.env
      });
      await this.logMessage(formatting.getStatus(CheckStatus.Failure, message));
      throw `SharePoint Framework v${spfxVersion} is not supported in ${SharePointVersion[sp]}`;
    }
  }

  private async checkNodeVersion(prerequisites: SpfxVersionPrerequisites): Promise<void> {
    const nodeVersion: string = this.getNodeVersion();
    await this.checkStatus('Node', nodeVersion, prerequisites.node);
  }

  private async checkSharePointFrameworkVersion(spfxVersionRequested: string): Promise<void> {
    let spfxVersionDetected = await this.getSPFxVersionFromYoRcFile();
    if (!spfxVersionDetected) {
      spfxVersionDetected = await this.getPackageVersion('@microsoft/generator-sharepoint', PackageSearchMode.GlobalOnly, HandlePromise.Continue);
    }
    const versionCheck: VersionCheck = {
      range: spfxVersionRequested,
      fix: `npm i -g @microsoft/generator-sharepoint@${spfxVersionRequested}`
    };
    if (spfxVersionDetected) {
      await this.checkStatus(`SharePoint Framework`, spfxVersionDetected, versionCheck);
    }
    else {
      const message = `SharePoint Framework v${spfxVersionRequested} not found`;
      this.resultsObject.push({
        check: 'SharePoint Framework',
        passed: false,
        version: spfxVersionRequested,
        message: message,
        fix: versionCheck.fix
      });
      await this.logMessage(formatting.getStatus(CheckStatus.Failure, message));
    }
  }

  private async checkYo(prerequisites: SpfxVersionPrerequisites): Promise<void> {
    const yoVersion: string = await this.getPackageVersion('yo', PackageSearchMode.GlobalOnly, HandlePromise.Continue);
    if (yoVersion) {
      await this.checkStatus('yo', yoVersion, prerequisites.yo);
    }
    else {
      const message = 'yo not found';
      this.resultsObject.push({
        check: 'yo',
        passed: false,
        message: message,
        fix: prerequisites.yo.fix
      });
      await this.logMessage(formatting.getStatus(CheckStatus.Failure, message));
    }
  }

  private async checkGulpCli(prerequisites: SpfxVersionPrerequisites): Promise<void> {
    const gulpCliVersion: string = await this.getPackageVersion('gulp-cli', PackageSearchMode.GlobalOnly, HandlePromise.Continue);
    if (gulpCliVersion) {
      await this.checkStatus('gulp-cli', gulpCliVersion, prerequisites.gulpCli);
    }
    else {
      const message = 'gulp-cli not found';
      this.resultsObject.push({
        check: 'gulp-cli',
        passed: false,
        message: message,
        fix: prerequisites.gulpCli.fix
      });
      await this.logMessage(formatting.getStatus(CheckStatus.Failure, message));
    }
  }

  private async checkGulp(): Promise<void> {
    const gulpVersion: string = await this.getPackageVersion('gulp', PackageSearchMode.GlobalOnly, HandlePromise.Continue);
    if (gulpVersion) {
      const message = 'gulp should be removed';
      const fix = 'npm un -g gulp';
      this.resultsObject.push({
        check: 'gulp',
        passed: false,
        version: gulpVersion,
        message: message,
        fix: fix
      });
      await this.logMessage(formatting.getStatus(CheckStatus.Failure, message));
    }
  }

  private async checkTypeScript(): Promise<void> {
    const typeScriptVersion: string = await this.getPackageVersion('typescript', PackageSearchMode.LocalOnly, HandlePromise.Continue);
    if (typeScriptVersion) {
      const fix = 'npm un typescript';
      const message = `typescript v${typeScriptVersion} installed in the project`;
      this.resultsObject.push({
        check: 'typescript',
        passed: false,
        message: message,
        fix: fix
      });
      await this.logMessage(formatting.getStatus(CheckStatus.Failure, message));
    }
    else {
      const message = 'bundled typescript used';
      this.resultsObject.push({
        check: 'typescript',
        passed: true,
        message: message
      });
      await this.logMessage(formatting.getStatus(CheckStatus.Success, message));
    }
  }

  private spVersionStringToEnum(sp: string): SharePointVersion | undefined {
    return (<any>SharePointVersion)[sp.toUpperCase()];
  }

  private async getSPFxVersionFromYoRcFile(): Promise<string | undefined> {
    if (this.projectRootPath !== null) {
      const spfxVersion = this.getProjectVersion();
      if (spfxVersion) {
        if (this.debug) {
          await this.logger.logToStderr(`SPFx version retrieved from .yo-rc.json file. Retrieved version: ${spfxVersion}`);
        }
        return spfxVersion;
      }
    }
    return undefined;
  }

  private async getSharePointFrameworkVersion(): Promise<string> {
    let spfxVersion = await this.getSPFxVersionFromYoRcFile();
    if (spfxVersion) {
      return spfxVersion;
    }
    try {
      spfxVersion = await this.getPackageVersion('@microsoft/sp-core-library', PackageSearchMode.LocalOnly, HandlePromise.Fail);
      if (this.debug) {
        await this.logger.logToStderr(`Found @microsoft/sp-core-library@${spfxVersion}`);
      }
      return spfxVersion;
    }
    catch {
      if (this.debug) {
        await this.logger.logToStderr(`@microsoft/sp-core-library not found. Search for @microsoft/generator-sharepoint local or global...`);
      }

      try {
        return await this.getPackageVersion('@microsoft/generator-sharepoint', PackageSearchMode.LocalAndGlobal, HandlePromise.Fail);
      }
      catch (error: any) {
        if (this.debug) {
          await this.logger.logToStderr('@microsoft/generator-sharepoint not found');
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

  private async getPackageVersion(packageName: string, searchMode: PackageSearchMode, handlePromise: HandlePromise): Promise<string> {
    const args: string[] = ['ls', packageName, '--depth=0', '--json'];
    if (searchMode === PackageSearchMode.GlobalOnly) {
      args.push('-g');
    }

    let version: string;
    try {
      version = await this.getPackageVersionFromNpm(args);
    }
    catch {
      if (searchMode === PackageSearchMode.LocalAndGlobal) {
        args.push('-g');
        version = await this.getPackageVersionFromNpm(args);
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

  private getPackageVersionFromNpm(args: string[]): Promise<string> {
    return new Promise<string>(async (resolve: (version: string) => void, reject: (error: string) => void): Promise<void> => {
      const packageName: string = args[1];

      if (this.debug) {
        await this.logger.logToStderr(`Executing npm: ${args.join(' ')}...`);
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

  private async checkStatus(what: string, versionFound: string, versionCheck: VersionCheck): Promise<void> {
    if (versionFound) {
      if (satisfies(versionFound, versionCheck.range)) {
        const message = `${what} v${versionFound}`;
        this.resultsObject.push({
          check: what,
          passed: true,
          message: message,
          version: versionFound
        });
        await this.logMessage(formatting.getStatus(CheckStatus.Success, message));
      }
      else {
        const message = `${what} v${versionFound} found, v${versionCheck.range} required`;
        this.resultsObject.push({
          check: what,
          passed: false,
          version: versionFound,
          message: message,
          fix: versionCheck.fix
        });
        await this.logMessage(formatting.getStatus(CheckStatus.Failure, message));
      }
    }
  }
}

export default new SpfxDoctorCommand();
