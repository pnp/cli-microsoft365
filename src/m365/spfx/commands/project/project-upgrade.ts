import commands from '../../commands';
import Command, {
  CommandOption, CommandError, CommandAction, CommandValidate
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import * as path from 'path';
import * as fs from 'fs';
import { Finding, Hash, Dictionary } from './project-upgrade/';
import { Rule } from './project-upgrade/rules/Rule';
import * as os from 'os';
import { Project } from './model';
import { FindingToReport } from './project-upgrade/FindingToReport';
import { FN017001_MISC_npm_dedupe } from './project-upgrade/rules/FN017001_MISC_npm_dedupe';
import { ReportData, ReportDataModification } from './ReportData';
import { BaseProjectCommand } from './base-project-command';
import { FindingTour } from './project-upgrade/FindingTour';
import { FindingTourStep } from './project-upgrade/FindingTourStep';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  packageManager?: string;
  toVersion?: string;
  shell?: string;
}

class SpfxProjectUpgradeCommand extends BaseProjectCommand {
  public constructor() {
    super();
  }

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
    '1.11.0'
  ];
  private static packageCommands = {
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
  }

  private static copyCommands = {
    bash: {
      copyCommand: 'cp',
      copyDestinationParam: ' '
    },
    powershell: {
      copyCommand: 'Copy-Item',
      copyDestinationParam: ' -Destination '
    },
    cmd: {
      copyCommand: 'copy',
      copyDestinationParam: ' '
    }
  }

  private static createDirectoryCommands = {
    bash: {
      createDirectoryCommand: 'mkdir',
      createDirectoryPathParam: ' ',
      createDirectoryNameParam: '/',
      createDirectoryItemTypeParam: '',
    },
    powershell: {
      createDirectoryCommand: 'New-Item',
      createDirectoryPathParam: ' -Path "',
      createDirectoryNameParam: '" -Name "',
      createDirectoryItemTypeParam: '" -ItemType "directory"',
    },
    cmd: {
      createDirectoryCommand: 'mkdir',
      createDirectoryPathParam: ' "',
      createDirectoryNameParam: '\\',
      createDirectoryItemTypeParam: '"',
    }
  }

  private static addFileCommands = {
    bash: {
      addFileCommand: 'cat > [FILEPATH] << EOF [FILECONTENT]EOF',
    },
    powershell: {
      addFileCommand: `@"[FILECONTENT]"@ | Out-File -FilePath "[FILEPATH]"
      `,
    },
    cmd: {
      addFileCommand: `echo [FILECONTENT] > "[FILEPATH]"
      `,
    }
  }

  private static removeFileCommands = {
    bash: {
      removeFileCommand: 'rm',
    },
    powershell: {
      removeFileCommand: 'Remove-Item',
    },
    cmd: {
      removeFileCommand: 'del',
    }
  }

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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.toVersion = args.options.toVersion || this.supportedVersions[this.supportedVersions.length - 1];
    telemetryProps.packageManager = args.options.packageManager || 'npm';
    telemetryProps.shell = args.options.shell || 'bash';
    return telemetryProps;
  }

  public action(): CommandAction {
    const cmd: Command = this;
    return function (this: CommandInstance, args: CommandArgs, cb: (err?: any) => void) {
      args = (cmd as any).processArgs(args);
      (cmd as any).initAction(args, this);
      cmd.commandAction(this, args, cb);
    }
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    this.projectRootPath = this.getProjectRoot(process.cwd());
    if (this.projectRootPath === null) {
      cb(new CommandError(`Couldn't find project root folder`, SpfxProjectUpgradeCommand.ERROR_NO_PROJECT_ROOT_FOLDER));
      return;
    }

    this.toVersion = args.options.toVersion ? args.options.toVersion : this.supportedVersions[this.supportedVersions.length - 1];
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
      cb(new CommandError(`CLI for Microsoft 365 doesn't support upgrading projects build on SharePoint Framework v${this.projectVersion}`, SpfxProjectUpgradeCommand.ERROR_UNSUPPORTED_FROM_VERSION));
      return;
    }

    const posTo: number = this.supportedVersions.indexOf(this.toVersion);
    if (pos > posTo) {
      cb(new CommandError('You cannot downgrade a project', SpfxProjectUpgradeCommand.ERROR_NO_DOWNGRADE));
      return;
    }

    if (pos === posTo) {
      cmd.log(`Project doesn't need to be upgraded`);
      cb();
      return;
    }

    if (this.verbose) {
      cmd.log('Collecting project...');
    }
    const project: Project = this.getProject(this.projectRootPath);

    if (this.debug) {
      cmd.log('Collected project');
      cmd.log(project);
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
    catch (e) {
      cb(new CommandError(e.message));
      return;
    }
    if (this.packageManager === 'npm') {
      const npmDedupeRule: Rule = new FN017001_MISC_npm_dedupe();
      npmDedupeRule.visit(project, this.allFindings);
    }

    // dedupe
    const findings: Finding[] = this.allFindings.filter((f: Finding, i: number, allFindings: Finding[]) => {
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
        f.resolution = f.resolution.replace('uninstallDev', this.getPackageManagerCommand('uninstallDev'));
        return;
      }
      if (f.resolution.startsWith('installDev')) {
        f.resolution = f.resolution.replace('installDev', this.getPackageManagerCommand('installDev'));
        return;
      }
      if (f.resolution.startsWith('uninstall')) {
        f.resolution = f.resolution.replace('uninstall', this.getPackageManagerCommand('uninstall'));
        return;
      }
      if (f.resolution.startsWith('install')) {
        f.resolution = f.resolution.replace('install', this.getPackageManagerCommand('install'));
        return;
      }

      // copy support for multiple shells
      if (f.resolution.startsWith('copy_cmd')) {
        f.resolution = f.resolution.replace('copy_cmd', this.getCopyCommand('copyCommand'));
        f.resolution = f.resolution.replace('DestinationParam', this.getCopyCommand('copyDestinationParam'));
        return;
      }
      // createdir support for multiple shells
      if (f.resolution.startsWith('create_dir_cmd')) {
        f.resolution = f.resolution.replace('create_dir_cmd', this.getDirectoryCommand('createDirectoryCommand'));
        f.resolution = f.resolution.replace('NameParam', this.getDirectoryCommand('createDirectoryNameParam'));
        f.resolution = f.resolution.replace('PathParam', this.getDirectoryCommand('createDirectoryPathParam'));
        f.resolution = f.resolution.replace('ItemTypeParam', this.getDirectoryCommand('createDirectoryItemTypeParam'));
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

        f.resolution = this.getAddCommand('addFileCommand');
        f.resolution = f.resolution.replace('[FILECONTENT]', fileContent)
        f.resolution = f.resolution.replace('[FILEPATH]', filePath);
        f.resolution = f.resolution.replace('[BEFOREPATH]', ' ');
        f.resolution = f.resolution.replace('[AFTERPATH]', ' ');
        f.resolution = f.resolution.replace('[BEFORECONTENT]', ' ');
        f.resolution = f.resolution.replace('[AFTERCONTENT]', ' ');
        return;
      }
      // 'Remove' support for multiple shells
      if (f.resolution.startsWith('remove_cmd')) {
        f.resolution = f.resolution.replace('remove_cmd', this.getRemoveCommand('removeFileCommand'));
        return;
      }
    });

    switch (args.options.output) {
      case 'json':
        cmd.log(findingsToReport);
        break;
      case 'tour':
        this.writeReportTourFolder(this.getTourReport(findingsToReport, project), cmd, args.options);
        break;
      case 'md':
        cmd.log(this.getMdReport(findingsToReport));       
        break;
      default:
        cmd.log(this.getTextReport(findingsToReport));
    }

    cb();
  }

  private writeReportTourFolder(findingsToReport: any, cmd: CommandInstance, options: Options): void {
    const toursFolder: string = path.join(this.projectRootPath as string, '.tours');

    if (!fs.existsSync(toursFolder)) {
      fs.mkdirSync(toursFolder, { recursive: false });
    }
    const tourFilePath = path.join(this.projectRootPath as string, '.tours', 'upgrade.tour');
    fs.writeFileSync(path.resolve(tourFilePath),findingsToReport, 'utf-8');
  }

  private getTextReport(findings: FindingToReport[]): string {
    const reportData: ReportData = this.getReportData(findings);
    const s: string[] = [
      'Execute in ' + this.shell, os.EOL,
      '-----------------------', os.EOL,
      (reportData.packageManagerCommands
        .concat(reportData.commandsToExecute
          .filter((command) =>
            command.indexOf(this.getPackageManagerCommand('install')) === -1 &&
            command.indexOf(this.getPackageManagerCommand('installDev')) === -1 &&
            command.indexOf(this.getPackageManagerCommand('uninstall')) === -1 &&
            command.indexOf(this.getPackageManagerCommand('uninstallDev')) === -1))).join(os.EOL), os.EOL,
      os.EOL,
      Object.keys(reportData.modificationPerFile).map(file => {
        return [
          file, os.EOL,
          '-'.repeat(file.length), os.EOL,
          reportData.modificationPerFile[file].map((m: ReportDataModification) => `${m.description}:${os.EOL}${m.modification}${os.EOL}`).join(os.EOL), os.EOL,
        ].join('');
      }).join(os.EOL),
      os.EOL,
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
          resolution = `In file [${f.file}](${f.file}) update the code as follows:

\`\`\`${f.resolutionType}
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
            command.indexOf(this.getPackageManagerCommand('install')) === -1 &&
            command.indexOf(this.getPackageManagerCommand('installDev')) === -1 &&
            command.indexOf(this.getPackageManagerCommand('uninstall')) === -1 &&
            command.indexOf(this.getPackageManagerCommand('uninstallDev')) === -1))).join(os.EOL), os.EOL,
      '```', os.EOL,
      os.EOL,
      '### Modify files', os.EOL,
      os.EOL,
      Object.keys(reportData.modificationPerFile).map(file => {
        return [
          `#### [${file}](${file})`, os.EOL,
          os.EOL,
          reportData.modificationPerFile[file].map((m: ReportDataModification) => `${m.description}:${os.EOL}${os.EOL}\`\`\`${reportData.modificationTypePerFile[file]}${os.EOL}${m.modification}${os.EOL}\`\`\``).join(os.EOL + os.EOL), os.EOL,
        ].join('');
      }).join(os.EOL),
      os.EOL,
    ];

    return s.join('').trim();
  }

  private getTourReport(findings: FindingToReport[], project: Project): string {
    const tourFindings: FindingTour = {
      title: `Upgrade project ${path.basename(this.projectRootPath as string)} to v${this.toVersion}`,
      steps: []
    };

    findings.forEach(f => {
      const lineNumber: number = f.position && f.position.line ? f.position.line : this.getLineToModify(f, project.path);

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
        title: `${sev}: ${f.title}`,
        description: `### ${sev}\r\n\r\n${f.description}\r\n\r\n${resolution}`,
        line: lineNumber
      };

      // Point to a directory if there is no file
      if (file !== undefined) {
        step.file = file;
      } else {
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

  private getLineToModify(finding: FindingToReport, rootPath: string): number {
    const filePath: string = path.resolve(path.join(rootPath, finding.file));

    // Don't cause an issue if the file isn't there
    if (!fs.existsSync(filePath)) {
      return 1;
    }

    // Read the file content
    const fileContent: string = fs.readFileSync(filePath, 'utf-8');

    // Try to find the line this relates to
    const lines: string[] = fileContent.split('\n'); // os.EOL doesn't work here

    const lineIndex: number = lines.findIndex((line: string) => line.indexOf(finding.title) > -1);
    if (lineIndex < 1) {
      return 1;
    }
    else {
      // Lines are 1-based in files
      return lineIndex + 1;
    }
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
          this.mapPackageManagerCommand(f.resolution, packagesDevExact,
            packagesDepExact, packagesDepUn, packagesDevUn);
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

    const packageManagerCommands: string[] = this.reducePackageManagerCommand(
      packagesDepExact, packagesDevExact, packagesDepUn, packagesDevUn);

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
    }
  }

  private mapPackageManagerCommand(command: string, packagesDevExact: string[],
    packagesDepExact: string[], packagesDepUn: string[], packagesDevUn: string[]): void {
    // matches must be in this particular order to avoid false matches, eg.
    // uninstallDev contains install
    if (command.startsWith(`${this.getPackageManagerCommand('uninstallDev')} `)) {
      packagesDevUn.push(command.replace(this.getPackageManagerCommand('uninstallDev'), '').trim());
      return;
    }
    if (command.startsWith(`${this.getPackageManagerCommand('installDev')} `)) {
      packagesDevExact.push(command.replace(this.getPackageManagerCommand('installDev'), '').trim());
      return;
    }
    if (command.startsWith(`${this.getPackageManagerCommand('uninstall')} `)) {
      packagesDepUn.push(command.replace(this.getPackageManagerCommand('uninstall'), '').trim());
      return;
    }
    if (command.startsWith(`${this.getPackageManagerCommand('install')} `)) {
      packagesDepExact.push(command.replace(this.getPackageManagerCommand('install'), '').trim());
    }
  }

  private reducePackageManagerCommand(packagesDepExact: string[], packagesDevExact: string[],
    packagesDepUn: string[], packagesDevUn: string[]): string[] {
    const commandsToExecute: string[] = [];

    if (packagesDepExact.length > 0) {
      commandsToExecute.push(`${this.getPackageManagerCommand('install')} ${packagesDepExact.join(' ')}`);
    }

    if (packagesDevExact.length > 0) {
      commandsToExecute.push(`${this.getPackageManagerCommand('installDev')} ${packagesDevExact.join(' ')}`);
    }

    if (packagesDepUn.length > 0) {
      commandsToExecute.push(`${this.getPackageManagerCommand('uninstall')} ${packagesDepUn.join(' ')}`);
    }

    if (packagesDevUn.length > 0) {
      commandsToExecute.push(`${this.getPackageManagerCommand('uninstallDev')} ${packagesDevUn.join(' ')}`);
    }

    return commandsToExecute;
  }

  private getPackageManagerCommand(command: string): string {
    return (SpfxProjectUpgradeCommand.packageCommands as any)[this.packageManager][command];
  }

  private getCopyCommand(command: string): string {
    return (SpfxProjectUpgradeCommand.copyCommands as any)[this.shell][command];
  }

  private getDirectoryCommand(command: string): string {
    return (SpfxProjectUpgradeCommand.createDirectoryCommands as any)[this.shell][command];
  }

  private getAddCommand(command: string): string {
    return (SpfxProjectUpgradeCommand.addFileCommands as any)[this.shell][command];
  }

  private getRemoveCommand(command: string): string {
    return (SpfxProjectUpgradeCommand.removeFileCommands as any)[this.shell][command];
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-v, --toVersion [toVersion]',
        description: 'The version of SharePoint Framework to which upgrade the project'
      },
      {
        option: '--packageManager [packageManager]',
        description: 'The package manager you use. Supported managers npm|pnpm|yarn. Default npm',
        autocomplete: ['npm', 'pnpm', 'yarn']
      },
      {
        option: '--shell [shell]',
        description: 'The shell you use. Supported shells bash|powershell|cmd. Default bash',
        autocomplete: ['bash', 'powershell', 'cmd']
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    parentOptions.forEach(o => {
      if (o.option.indexOf('--output') > -1) {
        o.description = 'Output type. json|text|md|tour. Default text';
        o.autocomplete = ['json', 'text', 'md', 'tour'];
      }
    })
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (args.options.packageManager) {
        if (['npm', 'pnpm', 'yarn'].indexOf(args.options.packageManager) < 0) {
          return `${args.options.packageManager} is not a supported package manager. Supported package managers are npm, pnpm and yarn`;
        }
      }

      if (args.options.shell) {
        if (['bash', 'powershell', 'cmd'].indexOf(args.options.shell) < 0) {
          return `${args.options.shell} is not a supported shell. Supported shells are bash, powershell and cmd`;
        }
      }     

      return true;
    };
  }

  public commandHelp(args: any, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.PROJECT_UPGRADE).helpInformation());
    log(
      `   ${chalk.yellow('Important:')} Run this command in the folder where the project
    that you want to upgrade is located. This command doesn't change your
    project files.
      
  Remarks:

    The ${this.name} command helps you upgrade your SharePoint Framework
    project to the specified version. If no version is specified, the command
    will upgrade to the latest version of the SharePoint Framework it supports
    (v1.11.0).

    This command doesn't change your project files. Instead, it gives you
    a report with all steps necessary to upgrade your project to the specified
    version of the SharePoint Framework. Changing project files is error-prone,
    especially when it comes to updating your solution's code. This is why at
    this moment, this command produces a report that you can use yourself to
    perform the necessary updates and verify that everything is working as
    expected.

    Using this command you can upgrade SharePoint Framework projects built using
    versions: 1.0.0, 1.0.1, 1.0.2, 1.1.0, 1.1.1, 1.1.3, 1.2.0, 1.3.0, 1.3.1,
    1.3.2, 1.3.4, 1.4.0, 1.4.1, 1.5.0, 1.5.1, 1.6.0, 1.7.0, 1.7.1, 1.8.0,
    1.8.1, 1.8.2, 1.9.1 and 1.10.0.

  Examples:
  
    Get instructions to upgrade the current SharePoint Framework project to
    SharePoint Framework version 1.5.0 and save the findings in a Markdown file
      ${this.name} --toVersion 1.5.0 --output md > "upgrade-report.md"

    Get instructions to Upgrade the current SharePoint Framework project to
    SharePoint Framework version 1.5.0 and show the summary of the findings
    in the shell
      ${this.name} --toVersion 1.5.0

    Get instructions to upgrade the current SharePoint Framework project to the
    latest SharePoint Framework version supported by the CLI for Microsoft 365 using
    pnpm
      ${this.name} --packageManager pnpm

    Get instructions to upgrade the current SharePoint Framework project to the
    latest SharePoint Framework version supported by the CLI for Microsoft 365
      ${this.name}

    Get instructions to upgrade the current SharePoint Framework project to the
    latest SharePoint Framework version supported by the CLI for Microsoft 365 using
    PowerShell
      ${this.name} --shell powershell

    Get instructions to upgrade the current SharePoint Framework project to
    the latest version of SharePoint Framework and save the findings in a 
    CodeTour file
        ${this.name} --output tour
`);
  }
}

module.exports = new SpfxProjectUpgradeCommand();