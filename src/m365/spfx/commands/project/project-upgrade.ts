import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
// uncomment to support upgrading to preview releases
import { prerelease } from 'semver';
import { Logger } from '../../../../cli';
import { CommandError } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { Dictionary, fsUtil, Hash, packageManager } from '../../../../utils';
import commands from '../../commands';
import { BaseProjectCommand } from './base-project-command';
import { Project } from './project-model';
import { FN017001_MISC_npm_dedupe } from './project-upgrade/rules/FN017001_MISC_npm_dedupe';
import { Finding, FindingToReport, FindingTour, FindingTourStep } from './report-model';
import { ReportData, ReportDataModification } from './report-model/ReportData';
import { Rule } from './Rule';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  packageManager?: string;
  preview?: boolean;
  toVersion?: string;
  shell?: string;
}

class SpfxProjectUpgradeCommand extends BaseProjectCommand {
  private static packageManagers: string[] = ['npm', 'pnpm', 'yarn'];
  private static shells: string[] = ['bash', 'powershell', 'cmd'];

  private projectVersion: string | undefined;
  private toVersion: string = '';
  private packageManager: string = 'npm';
  private shell: string = 'bash';
  private allFindings: Finding[] = [];
  private supportedVersions: string[] = [
    '1.0.0',
    '1.0.1',
    '1.0.2',
    '1.1.0',
    '1.1.1',
    '1.1.3',
    '1.2.0',
    '1.3.0',
    '1.3.1',
    '1.3.2',
    '1.3.4',
    '1.4.0',
    '1.4.1',
    '1.5.0',
    '1.5.1',
    '1.6.0',
    '1.7.0',
    '1.7.1',
    '1.8.0',
    '1.8.1',
    '1.8.2',
    '1.9.1',
    '1.10.0',
    '1.11.0',
    '1.12.0',
    '1.12.1',
    '1.13.0',
    '1.13.1',
    '1.14.0',
    '1.15.0',
    '1.15.2',
    '1.16.0-beta.1'
  ];

  public static ERROR_NO_PROJECT_ROOT_FOLDER: number = 1;
  public static ERROR_UNSUPPORTED_TO_VERSION: number = 2;
  public static ERROR_NO_VERSION: number = 3;
  public static ERROR_UNSUPPORTED_FROM_VERSION: number = 4;
  public static ERROR_NO_DOWNGRADE: number = 5;

  public get name(): string {
    return commands.PROJECT_UPGRADE;
  }

  public get description(): string {
    return 'Upgrades SharePoint Framework project to the specified version';
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
        toVersion: args.options.toVersion || this.supportedVersions[this.supportedVersions.length - 1],
        packageManager: args.options.packageManager || 'npm',
        shell: args.options.shell || 'bash',
        preview: args.options.preview
      });
      // uncomment to support upgrading to preview releases
      if (prerelease(this.telemetryProperties.toVersion) && !args.options.preview) {
        this.telemetryProperties.toVersion = this.supportedVersions[this.supportedVersions.length - 2];
      }
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-v, --toVersion [toVersion]'
      },
      {
        option: '--packageManager [packageManager]',
        autocomplete: SpfxProjectUpgradeCommand.packageManagers
      },
      {
        option: '--shell [shell]',
        autocomplete: SpfxProjectUpgradeCommand.shells
      },
      {
        option: '--preview'
      }
    );

    this.options.forEach(o => {
      if (o.option.indexOf('--output') > -1) {
        o.autocomplete = ['json', 'text', 'md', 'tour'];
      }
    });
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.packageManager) {
          if (SpfxProjectUpgradeCommand.packageManagers.indexOf(args.options.packageManager) < 0) {
            return `${args.options.packageManager} is not a supported package manager. Supported package managers are ${SpfxProjectUpgradeCommand.packageManagers.join(', ')}`;
          }
        }
    
        if (args.options.shell) {
          if (SpfxProjectUpgradeCommand.shells.indexOf(args.options.shell) < 0) {
            return `${args.options.shell} is not a supported shell. Supported shells are ${SpfxProjectUpgradeCommand.shells.join(', ')}`;
          }
        }
    
        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    this.projectRootPath = this.getProjectRoot(process.cwd());
    if (this.projectRootPath === null) {
      cb(new CommandError(`Couldn't find project root folder`, SpfxProjectUpgradeCommand.ERROR_NO_PROJECT_ROOT_FOLDER));
      return;
    }

    this.toVersion = args.options.toVersion ? args.options.toVersion : this.supportedVersions[this.supportedVersions.length - 1];
    // uncomment to support upgrading to preview releases
    if (!args.options.toVersion &&
      !args.options.preview &&
      prerelease(this.toVersion)) {
      // no version and no preview specified while the current version to
      // upgrade to is a prerelease so let's grab the first non-preview version
      // since we're supporting only one preview version, it's sufficient for
      // us to take second to last version
      this.toVersion = this.supportedVersions[this.supportedVersions.length - 2];
    }
    this.packageManager = args.options.packageManager || 'npm';
    this.shell = args.options.shell || 'bash';

    if (this.supportedVersions.indexOf(this.toVersion) < 0) {
      cb(new CommandError(`CLI for Microsoft 365 doesn't support upgrading SharePoint Framework projects to version ${this.toVersion}. Supported versions are ${this.supportedVersions.join(', ')}`, SpfxProjectUpgradeCommand.ERROR_UNSUPPORTED_TO_VERSION));
      return;
    }

    this.projectVersion = this.getProjectVersion();
    if (!this.projectVersion) {
      cb(new CommandError(`Unable to determine the version of the current SharePoint Framework project`, SpfxProjectUpgradeCommand.ERROR_NO_VERSION));
      return;
    }

    const pos: number = this.supportedVersions.indexOf(this.projectVersion);
    if (pos < 0) {
      cb(new CommandError(`CLI for Microsoft 365 doesn't support upgrading projects built using SharePoint Framework v${this.projectVersion}`, SpfxProjectUpgradeCommand.ERROR_UNSUPPORTED_FROM_VERSION));
      return;
    }

    const posTo: number = this.supportedVersions.indexOf(this.toVersion);
    if (pos > posTo) {
      cb(new CommandError('You cannot downgrade a project', SpfxProjectUpgradeCommand.ERROR_NO_DOWNGRADE));
      return;
    }

    if (pos === posTo) {
      logger.log(`Project doesn't need to be upgraded`);
      cb();
      return;
    }

    if (this.verbose) {
      logger.logToStderr('Collecting project...');
    }
    const project: Project = this.getProject(this.projectRootPath);

    if (this.debug) {
      logger.logToStderr('Collected project');
      logger.logToStderr(project);
    }

    // reverse the list of versions to upgrade to, so that most recent findings
    // will end up on top already. Saves us reversing a larger array later
    const versionsToUpgradeTo: string[] = this.supportedVersions.slice(pos + 1, posTo + 1).reverse();
    try {
      versionsToUpgradeTo.forEach(v => {
        const rules: Rule[] = require(`./project-upgrade/upgrade-${v}`);
        rules.forEach(r => {
          r.visit(project, this.allFindings);
        });
      });
    }
    catch (e: any) {
      cb(new CommandError(e.message));
      return;
    }
    if (this.packageManager === 'npm') {
      const npmDedupeRule: Rule = new FN017001_MISC_npm_dedupe();
      npmDedupeRule.visit(project, this.allFindings);
    }

    // dedupe
    const findings: Finding[] = this.allFindings.filter((f: Finding, i: number) => {
      const firstFindingPos: number = this.allFindings.findIndex(f1 => f1.id === f.id);
      return i === firstFindingPos;
    });

    // remove superseded findings
    findings
      // get findings that supersede other findings
      .filter(f => f.supersedes.length > 0)
      .forEach(f => {
        f.supersedes.forEach(s => {
          // find the superseded finding
          const i: number = findings.findIndex(f1 => f1.id === s);
          if (i > -1) {
            // ...and remove it from findings
            findings.splice(i, 1);
          }
        });
      });

    // remove findings without title
    findings
      .forEach(f => {
        if (!f.title) {
          // find the finding
          const i: number = findings.findIndex(f1 => f1.id === f.id);
          if (i > -1) {
            // ...and remove it from findings
            findings.splice(i, 1);
          }
        }
      });

    // flatten
    const findingsToReport: FindingToReport[] = ([] as FindingToReport[]).concat.apply([], findings.map(f => {
      return f.occurrences.map(o => {
        return {
          description: f.description,
          id: f.id,
          file: o.file,
          position: o.position,
          resolution: o.resolution,
          resolutionType: f.resolutionType,
          severity: f.severity,
          title: f.title
        };
      });
    }));

    // replace package operation tokens with command for the specific package manager
    findingsToReport.forEach(f => {
      // matches must be in this particular order to avoid false matches, eg.
      // uninstallDev contains install
      if (f.resolution.startsWith('uninstallDev')) {
        f.resolution = f.resolution.replace('uninstallDev', packageManager.getPackageManagerCommand('uninstallDev', this.packageManager));
        return;
      }
      if (f.resolution.startsWith('installDev')) {
        f.resolution = f.resolution.replace('installDev', packageManager.getPackageManagerCommand('installDev', this.packageManager));
        return;
      }
      if (f.resolution.startsWith('uninstall')) {
        f.resolution = f.resolution.replace('uninstall', packageManager.getPackageManagerCommand('uninstall', this.packageManager));
        return;
      }
      if (f.resolution.startsWith('install')) {
        f.resolution = f.resolution.replace('install', packageManager.getPackageManagerCommand('install', this.packageManager));
        return;
      }

      // copy support for multiple shells
      if (f.resolution.startsWith('copy_cmd')) {
        f.resolution = f.resolution.replace('copy_cmd', fsUtil.getCopyCommand('copyCommand', this.shell));
        f.resolution = f.resolution.replace('DestinationParam', fsUtil.getCopyCommand('copyDestinationParam', this.shell));
        return;
      }
      // createdir support for multiple shells
      if (f.resolution.startsWith('create_dir_cmd')) {
        f.resolution = f.resolution.replace('create_dir_cmd', fsUtil.getDirectoryCommand('createDirectoryCommand', this.shell));
        f.resolution = f.resolution.replace('NameParam', fsUtil.getDirectoryCommand('createDirectoryNameParam', this.shell));
        f.resolution = f.resolution.replace('PathParam', fsUtil.getDirectoryCommand('createDirectoryPathParam', this.shell));
        f.resolution = f.resolution.replace('ItemTypeParam', fsUtil.getDirectoryCommand('createDirectoryItemTypeParam', this.shell));
        return;
      }
      // 'Add' support for multiple shells
      if (f.resolution.startsWith('add_cmd')) {
        const pathStart: number = f.resolution.indexOf('[BEFOREPATH]') + '[BEFOREPATH]'.length;
        const pathEnd: number = f.resolution.indexOf('[AFTERPATH]');
        const filePath: string = f.resolution.substring(pathStart, pathEnd);

        const contentStart: number = f.resolution.indexOf('[BEFORECONTENT]') + '[BEFORECONTENT]'.length;
        const contentEnd: number = f.resolution.indexOf('[AFTERCONTENT]');
        const fileContent: string = f.resolution.substring(contentStart, contentEnd);

        f.resolution = fsUtil.getAddCommand('addFileCommand', this.shell);
        f.resolution = f.resolution.replace('[FILECONTENT]', fileContent);
        f.resolution = f.resolution.replace('[FILEPATH]', filePath);
        f.resolution = f.resolution.replace('[BEFOREPATH]', ' ');
        f.resolution = f.resolution.replace('[AFTERPATH]', ' ');
        f.resolution = f.resolution.replace('[BEFORECONTENT]', ' ');
        f.resolution = f.resolution.replace('[AFTERCONTENT]', ' ');
        return;
      }
      // 'Remove' support for multiple shells
      if (f.resolution.startsWith('remove_cmd')) {
        f.resolution = f.resolution.replace('remove_cmd', fsUtil.getRemoveCommand('removeFileCommand', this.shell));
        return;
      }
    });

    switch (args.options.output) {
      case 'text':
        logger.log(this.getTextReport(findingsToReport));
        break;
      case 'json':
        logger.log(findingsToReport);
        break;
      case 'tour':
        this.writeReportTourFolder(this.getTourReport(findingsToReport, project));
        break;
      case 'md':
        logger.log(this.getMdReport(findingsToReport));
        break;
      default:
        logger.log(findingsToReport);
    }

    cb();
  }

  private writeReportTourFolder(findingsToReport: any): void {
    const toursFolder: string = path.join(this.projectRootPath as string, '.tours');

    if (!fs.existsSync(toursFolder)) {
      fs.mkdirSync(toursFolder, { recursive: false });
    }

    const tourFilePath: string = path.join(this.projectRootPath as string, '.tours', 'upgrade.tour');
    fs.writeFileSync(path.resolve(tourFilePath), findingsToReport, 'utf-8');
  }

  private getTextReport(findings: FindingToReport[]): string {
    const reportData: ReportData = this.getReportData(findings);
    const s: string[] = [
      'Execute in ' + this.shell, os.EOL,
      '-----------------------', os.EOL,
      (reportData.packageManagerCommands
        .concat(reportData.commandsToExecute
          .filter((command) =>
            command.indexOf(packageManager.getPackageManagerCommand('install', this.packageManager)) === -1 &&
            command.indexOf(packageManager.getPackageManagerCommand('installDev', this.packageManager)) === -1 &&
            command.indexOf(packageManager.getPackageManagerCommand('uninstall', this.packageManager)) === -1 &&
            command.indexOf(packageManager.getPackageManagerCommand('uninstallDev', this.packageManager)) === -1))).join(os.EOL), os.EOL,
      os.EOL,
      Object.keys(reportData.modificationPerFile).map(file => {
        return [
          file, os.EOL,
          '-'.repeat(file.length), os.EOL,
          reportData.modificationPerFile[file].map((m: ReportDataModification) => `${m.description}:${os.EOL}${m.modification}${os.EOL}`).join(os.EOL), os.EOL
        ].join('');
      }).join(os.EOL),
      os.EOL
    ];

    return s.join('').trim();
  }

  private getMdReport(findings: FindingToReport[]): string {
    const findingsToReport: string[] = [];
    const reportData: ReportData = this.getReportData(findings);

    findings.forEach(f => {
      let resolution: string = '';
      switch (f.resolutionType) {
        case 'cmd':
          resolution = `Execute the following command:

\`\`\`sh
${f.resolution}
\`\`\`
`;
          break;
        case 'json':
        case 'js':
        case 'ts':
        case 'scss':
          resolution = `\`\`\`${f.resolutionType}
${f.resolution}
\`\`\`
`;
          break;
      }

      findingsToReport.push(
        `### ${f.id} ${f.title} | ${f.severity}`, os.EOL,
        os.EOL,
        f.description, os.EOL,
        os.EOL,
        resolution,
        os.EOL,
        `File: [${f.file}${(f.position ? `:${f.position.line}:${f.position.character}` : '')}](${f.file})`, os.EOL,
        os.EOL
      );
    });

    const s: string[] = [
      `# Upgrade project ${path.basename(this.projectRootPath as string)} to v${this.toVersion}`, os.EOL,
      os.EOL,
      `Date: ${(new Date().toLocaleDateString())}`, os.EOL,
      os.EOL,
      '## Findings', os.EOL,
      os.EOL,
      `Following is the list of steps required to upgrade your project to SharePoint Framework version ${this.toVersion}. [Summary](#Summary) of the modifications is included at the end of the report.`, os.EOL,
      os.EOL,
      findingsToReport.join(''),
      '## Summary', os.EOL,
      os.EOL,
      '### Execute script', os.EOL,
      os.EOL,
      '```sh', os.EOL,
      (reportData.packageManagerCommands
        .concat(reportData.commandsToExecute
          .filter((command) =>
            command.indexOf(packageManager.getPackageManagerCommand('install', this.packageManager)) === -1 &&
            command.indexOf(packageManager.getPackageManagerCommand('installDev', this.packageManager)) === -1 &&
            command.indexOf(packageManager.getPackageManagerCommand('uninstall', this.packageManager)) === -1 &&
            command.indexOf(packageManager.getPackageManagerCommand('uninstallDev', this.packageManager)) === -1))).join(os.EOL), os.EOL,
      '```', os.EOL,
      os.EOL,
      '### Modify files', os.EOL,
      os.EOL,
      Object.keys(reportData.modificationPerFile).map(file => {
        return [
          `#### [${file}](${file})`, os.EOL,
          os.EOL,
          reportData.modificationPerFile[file].map((m: ReportDataModification) => `${m.description}:${os.EOL}${os.EOL}\`\`\`${reportData.modificationTypePerFile[file]}${os.EOL}${m.modification}${os.EOL}\`\`\``).join(os.EOL + os.EOL), os.EOL
        ].join('');
      }).join(os.EOL),
      os.EOL
    ];

    return s.join('').trim();
  }

  private getTourReport(findings: FindingToReport[], project: Project): string {
    const tourFindings: FindingTour = {
      title: `Upgrade project ${path.basename(this.projectRootPath as string)} to v${this.toVersion}`,
      steps: []
    };

    findings.forEach(f => {
      const lineNumber: number = f.position && f.position.line ? f.position.line : 1;

      let resolution: string = '';
      switch (f.resolutionType) {
        case 'cmd':
          resolution = `Execute the following command:\r\n\r\n[\`${f.resolution}\`](command:codetour.sendTextToTerminal?["${f.resolution}"])`;
          break;
        case 'json':
        case 'js':
        case 'ts':
        case 'scss':
          resolution = `\r\n\`\`\`${f.resolutionType}\r\n${f.resolution}\r\n\`\`\``;
          break;
      }

      // Make severity uppercase for the markdown
      const sev: string = f.severity.toUpperCase();

      // Clean up the file name
      let file: string | undefined = fs.existsSync(path.join(project.path, f.file)) ? f.file : undefined;

      if (file !== undefined) {
        // CodeTour expects the files to be relative from root (i.e.: no './')
        file = file.replace(/\.\//g, '');

        // CodeTour also expects forward slashes as directory separators
        file = file.replace(/\\/g, '/');
      }

      // Create a tour step entry
      const step: FindingTourStep = {
        title: `${sev}: ${f.title} (${f.id})`,
        description: `### ${sev}\r\n\r\n${f.description}\r\n\r\n${resolution}`,
        line: lineNumber
      };

      // Point to a directory if there is no file
      if (file !== undefined) {
        step.file = file;
      }
      else {
        step.directory = "";
      }

      tourFindings.steps.push(step);
    });

    // Add the finale
    tourFindings.steps.push({
      file: ".tours/upgrade.tour",
      title: "RECOMMENDED: Delete tour",
      description: "### THAT'S IT!!!\r\nOnce you have tested that your upgrade is successful, you can delete the `.tour` folder and its contents. Otherwise, you'll be prompted to launch this CodeTour every time you open this project."
    });

    return JSON.stringify(tourFindings, null, 2);
  }

  private getReportData(findings: FindingToReport[]): ReportData {
    const commandsToExecute: string[] = [];
    const modificationPerFile: Dictionary<ReportDataModification[]> = {};
    const modificationTypePerFile: Hash = {};
    const packagesDevExact: string[] = [];
    const packagesDepExact: string[] = [];
    const packagesDepUn: string[] = [];
    const packagesDevUn: string[] = [];

    findings.forEach(f => {
      if (f.resolutionType === 'cmd') {
        if (f.resolution.indexOf('npm') > -1 ||
          f.resolution.indexOf('yarn') > -1) {
          packageManager.mapPackageManagerCommand({
            command: f.resolution,
            packagesDevExact,
            packagesDepExact,
            packagesDepUn,
            packagesDevUn,
            packageMgr: this.packageManager
          });
        }
        else {
          commandsToExecute.push(f.resolution);
        }
      }
      else {
        if (!modificationPerFile[f.file]) {
          modificationPerFile[f.file] = [];
        }
        if (!modificationTypePerFile[f.file]) {
          modificationTypePerFile[f.file] = f.resolutionType;
        }

        modificationPerFile[f.file].push({
          description: f.description,
          modification: f.resolution
        });
      }
    });

    const packageManagerCommands: string[] = packageManager.reducePackageManagerCommand({
      packagesDepExact,
      packagesDevExact,
      packagesDepUn,
      packagesDevUn,
      packageMgr: this.packageManager
    });

    if (this.packageManager === 'npm') {
      const dedupeFinding: FindingToReport[] = findings.filter(f => f.id === 'FN017001');
      if (dedupeFinding.length > 0) {
        packageManagerCommands.push(dedupeFinding[0].resolution);
      }
    }

    return {
      commandsToExecute: commandsToExecute,
      packageManagerCommands: packageManagerCommands,
      modificationPerFile: modificationPerFile,
      modificationTypePerFile: modificationTypePerFile
    };
  }
}

module.exports = new SpfxProjectUpgradeCommand();
