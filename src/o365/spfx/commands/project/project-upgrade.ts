import config from '../../../../config';
import commands from '../../commands';
import Command, {
  CommandOption, CommandError
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import * as path from 'path';
import * as fs from 'fs';
import { Finding, Utils, Hash, Dictionary } from './project-upgrade/';
import { Rule } from './project-upgrade/rules/Rule';
import { EOL } from 'os';
import { Project, Manifest, TsFile } from './project-upgrade/model';
import { FindingToReport } from './project-upgrade/FindingToReport';
import { FN017001_MISC_npm_dedupe } from './project-upgrade/rules/FN017001_MISC_npm_dedupe';
import { ReportData, ReportDataModification } from './ReportData';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  toVersion?: string;
}

class SpfxProjectUpgradeCommand extends Command {
  private projectVersion: string | undefined;
  private toVersion: string = '';
  private projectRootPath: string | null = null;
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
    '1.6.0'
  ];

  public static ERROR_NO_PROJECT_ROOT_FOLDER: number = 1;
  public static ERROR_UNSUPPORTED_TO_VERSION: number = 2;
  public static ERROR_NO_VERSION: number = 3;
  public static ERROR_UNSUPPORTED_FROM_VERSION: number = 4;
  public static ERROR_NO_DOWNGRADE: number = 5;
  public static ERROR_PROJECT_UP_TO_DATE: number = 6;

  public get name(): string {
    return commands.PROJECT_UPGRADE;
  }

  public get description(): string {
    return 'Upgrades SharePoint Framework project to the specified version';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.toVersion = args.options.toVersion || this.supportedVersions[this.supportedVersions.length - 1];
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    this.projectRootPath = this.getProjectRoot(process.cwd());
    if (this.projectRootPath === null) {
      cb(new CommandError(`Couldn't find project root folder`, SpfxProjectUpgradeCommand.ERROR_NO_PROJECT_ROOT_FOLDER));
      return;
    }

    this.toVersion = args.options.toVersion ? args.options.toVersion : this.supportedVersions[this.supportedVersions.length - 1];

    if (this.supportedVersions.indexOf(this.toVersion) < 0) {
      cb(new CommandError(`Office 365 CLI doesn't support upgrading SharePoint Framework projects to version ${this.toVersion}. Supported versions are ${this.supportedVersions.join(', ')}`, SpfxProjectUpgradeCommand.ERROR_UNSUPPORTED_TO_VERSION));
      return;
    }

    this.projectVersion = this.getProjectVersion();
    if (!this.projectVersion) {
      cb(new CommandError(`Unable to determine the version of the current SharePoint Framework project`, SpfxProjectUpgradeCommand.ERROR_NO_VERSION));
      return;
    }

    const pos: number = this.supportedVersions.indexOf(this.projectVersion);
    if (pos < 0) {
      cb(new CommandError(`Office 365 CLI doesn't support upgrading projects build on SharePoint Framework v${this.projectVersion}`, SpfxProjectUpgradeCommand.ERROR_UNSUPPORTED_FROM_VERSION));
      return;
    }

    const posTo: number = this.supportedVersions.indexOf(this.toVersion);
    if (pos > posTo) {
      cb(new CommandError('You cannot downgrade a project', SpfxProjectUpgradeCommand.ERROR_NO_DOWNGRADE));
      return;
    }

    if (pos === posTo) {
      cb(new CommandError('Project doesn\'t need to be upgraded', SpfxProjectUpgradeCommand.ERROR_PROJECT_UP_TO_DATE));
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
    versionsToUpgradeTo.forEach(v => {
      try {
        const rules: Rule[] = require(`./project-upgrade/upgrade-${v}`);
        rules.forEach(r => {
          r.visit(project, this.allFindings);
        });
      }
      catch (e) {
        cb(new CommandError(e));
        return;
      }
    });
    const npmDedupeRule: Rule = new FN017001_MISC_npm_dedupe();
    npmDedupeRule.visit(project, this.allFindings);

    // dedupe
    const findings: Finding[] = this.allFindings.filter((f: Finding, i: number, allFindings: Finding[]) => {
      const firstFindingPos: number = this.allFindings.findIndex(f1 => f1.id === f.id);
      return i === firstFindingPos;
    });

    // flatten
    const findingsToReport: FindingToReport[] = [].concat.apply([], findings.map(f => {
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

    switch (args.options.output) {
      case 'json':
        cmd.log(findingsToReport);
        break;
      case 'md':
        cmd.log(this.getMdReport(findingsToReport));
        break;
      default:
        cmd.log(this.getTextReport(findingsToReport));
    }

    cb();
  }

  private getProject(projectRootPath: string): Project {
    const project: Project = {
      path: projectRootPath
    };

    const configJsonPath: string = path.join(projectRootPath, 'config/config.json');
    if (fs.existsSync(configJsonPath)) {
      try {
        project.configJson = JSON.parse(Utils.removeSingleLineComments(fs.readFileSync(configJsonPath, 'utf-8')));
      }
      catch { }
    }

    const copyAssetsJsonPath: string = path.join(projectRootPath, 'config/copy-assets.json');
    if (fs.existsSync(copyAssetsJsonPath)) {
      try {
        project.copyAssetsJson = JSON.parse(Utils.removeSingleLineComments(fs.readFileSync(copyAssetsJsonPath, 'utf-8')));
      }
      catch { }
    }

    const deployAzureStorageJsonPath: string = path.join(projectRootPath, 'config/deploy-azure-storage.json');
    if (fs.existsSync(deployAzureStorageJsonPath)) {
      try {
        project.deployAzureStorageJson = JSON.parse(Utils.removeSingleLineComments(fs.readFileSync(deployAzureStorageJsonPath, 'utf-8')));
      }
      catch { }
    }

    const packageJsonPath: string = path.join(projectRootPath, 'package.json');
    if (fs.existsSync(packageJsonPath)) {
      try {
        project.packageJson = JSON.parse(Utils.removeSingleLineComments(fs.readFileSync(packageJsonPath, 'utf-8')));
      }
      catch { }
    }

    const packageSolutionJsonPath: string = path.join(projectRootPath, 'config/package-solution.json');
    if (fs.existsSync(packageSolutionJsonPath)) {
      try {
        project.packageSolutionJson = JSON.parse(Utils.removeSingleLineComments(fs.readFileSync(packageSolutionJsonPath, 'utf-8')));
      }
      catch { }
    }

    const serveJsonPath: string = path.join(projectRootPath, 'config/serve.json');
    if (fs.existsSync(serveJsonPath)) {
      try {
        project.serveJson = JSON.parse(Utils.removeSingleLineComments(fs.readFileSync(serveJsonPath, 'utf-8')));
      }
      catch { }
    }

    const tsConfigJsonPath: string = path.join(projectRootPath, 'tsconfig.json');
    if (fs.existsSync(tsConfigJsonPath)) {
      try {
        project.tsConfigJson = JSON.parse(Utils.removeSingleLineComments(fs.readFileSync(tsConfigJsonPath, 'utf-8')));
      }
      catch { }
    }

    const tsLintJsonPath: string = path.join(projectRootPath, 'config/tslint.json');
    if (fs.existsSync(tsLintJsonPath)) {
      try {
        project.tsLintJson = JSON.parse(Utils.removeSingleLineComments(fs.readFileSync(tsLintJsonPath, 'utf-8')));
      }
      catch { }
    }

    const writeManifestJsonPath: string = path.join(projectRootPath, 'config/write-manifests.json');
    if (fs.existsSync(writeManifestJsonPath)) {
      try {
        project.writeManifestsJson = JSON.parse(Utils.removeSingleLineComments(fs.readFileSync(writeManifestJsonPath, 'utf-8')));
      }
      catch { }
    }

    const yoRcJsonPath: string = path.join(projectRootPath, '.yo-rc.json');
    if (fs.existsSync(yoRcJsonPath)) {
      try {
        project.yoRcJson = JSON.parse(Utils.removeSingleLineComments(fs.readFileSync(yoRcJsonPath, 'utf-8')));
      }
      catch { }
    }

    const manifests: Manifest[] = [];
    const files: string[] = Utils.getAllFiles(path.join(projectRootPath, 'src')) as string[];
    files.forEach(f => {
      if (f.endsWith('.manifest.json')) {
        try {
          const manifestStr = Utils.removeSingleLineComments(fs.readFileSync(f, 'utf-8'));
          const manifest: Manifest = JSON.parse(manifestStr);
          manifest.path = f;
          manifests.push(manifest);
        }
        catch { }
      }
    });
    project.manifests = manifests;

    const gulpfileJsPath: string = path.join(projectRootPath, 'gulpfile.js');
    if (fs.existsSync(gulpfileJsPath)) {
      project.gulpfileJs = {
        src: fs.readFileSync(gulpfileJsPath, 'utf-8')
      };
    }

    project.vsCode = {};
    const vsCodeSettingsPath: string = path.join(projectRootPath, '.vscode', 'settings.json');
    if (fs.existsSync(vsCodeSettingsPath)) {
      try {
        project.vsCode.settingsJson = JSON.parse(Utils.removeSingleLineComments(fs.readFileSync(vsCodeSettingsPath, 'utf-8')));
      }
      catch { }
    }

    const vsCodeExtensionsPath: string = path.join(projectRootPath, '.vscode', 'extensions.json');
    if (fs.existsSync(vsCodeExtensionsPath)) {
      try {
        project.vsCode.extensionsJson = JSON.parse(Utils.removeSingleLineComments(fs.readFileSync(vsCodeExtensionsPath, 'utf-8')));
      }
      catch { }
    }

    const vsCodeLaunchPath: string = path.join(projectRootPath, '.vscode', 'launch.json');
    if (fs.existsSync(vsCodeLaunchPath)) {
      try {
        project.vsCode.launchJson = JSON.parse(Utils.removeSingleLineComments(fs.readFileSync(vsCodeLaunchPath, 'utf-8')));
      }
      catch { }
    }

    const srcFiles: string[] = SpfxProjectUpgradeCommand.readdirR(path.join(projectRootPath, 'src')) as string[];
    const tsFiles: string[] = srcFiles.filter(f => f.endsWith('.ts') || f.endsWith('.tsx'));
    project.tsFiles = tsFiles.map(f => new TsFile(f));

    return project;
  }

  private static readdirR(dir: string): string | string[] {
    return fs.statSync(dir).isDirectory()
      ? Array.prototype.concat(...fs.readdirSync(dir).map(f => SpfxProjectUpgradeCommand.readdirR(path.join(dir, f))))
      : dir;
  }

  private getTextReport(findings: FindingToReport[]): string {
    const reportData: ReportData = this.getReportData(findings);
    const s: string[] = [
      'Execute in command line', EOL,
      '-----------------------', EOL,
      (reportData.mainNpmCommands.concat(reportData.commandsToExecute.filter((command) => command.indexOf('npm i') === -1 && command.indexOf('npm un') === -1))).join(EOL), EOL,
      EOL,
      Object.keys(reportData.modificationPerFile).map(file => {
        return [
          file, EOL,
          '-'.repeat(file.length), EOL,
          reportData.modificationPerFile[file].map((m: ReportDataModification) => `${m.description}:${EOL}${m.modification}${EOL}`).join(EOL), EOL,
        ].join('');
      }).join(EOL),
      EOL,
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
          resolution = `In file [${f.file}](${f.file}) update the code as follows:

\`\`\`${f.resolutionType}
${f.resolution}
\`\`\`
`;
          break;
      }

      findingsToReport.push(
        `### ${f.id} ${f.title} | ${f.severity}`, EOL,
        EOL,
        f.description, EOL,
        EOL,
        resolution,
        EOL,
        `File: [${f.file}${(f.position ? `:${f.position.line}:${f.position.character}` : '')}](${f.file})`, EOL,
        EOL
      );
    });

    const s: string[] = [
      `# Upgrade project ${path.posix.basename(this.projectRootPath as string)} to v${this.toVersion}`, EOL,
      EOL,
      `Date: ${(new Date().toLocaleDateString())}`, EOL,
      EOL,
      '## Findings', EOL,
      EOL,
      `Following is the list of steps required to upgrade your project to SharePoint Framework version ${this.toVersion}. [Summary](#Summary) of the modifications is included at the end of the report.`, EOL,
      EOL,
      findingsToReport.join(''),
      '## Summary', EOL,
      EOL,
      '### Execute script', EOL,
      EOL,
      '```sh', EOL,
      (reportData.mainNpmCommands.concat(reportData.commandsToExecute.filter((command) => command.indexOf('npm i') === -1 && command.indexOf('npm un') === -1))).join(EOL), EOL,
      '```', EOL,
      EOL,
      '### Modify files', EOL,
      EOL,
      Object.keys(reportData.modificationPerFile).map(file => {
        return [
          `#### [${file}](${file})`, EOL,
          EOL,
          reportData.modificationPerFile[file].map((m: ReportDataModification) => `${m.description}:${EOL}${EOL}\`\`\`${reportData.modificationTypePerFile[file]}${EOL}${m.modification}${EOL}\`\`\``).join(EOL + EOL), EOL,
        ].join('');
      }).join(EOL),
      EOL,
    ];

    return s.join('').trim();
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
        if (f.resolution.indexOf('npm i') !== -1 ||
          f.resolution.indexOf('npm un') !== -1) {
          this.mapNpmCommand(f.resolution, packagesDevExact,
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

    const mainNpmCommands: string[] = this.reduceNpmCommand(
      packagesDepExact, packagesDevExact, packagesDepUn, packagesDevUn);

    return {
      commandsToExecute: commandsToExecute,
      mainNpmCommands: mainNpmCommands,
      modificationPerFile: modificationPerFile,
      modificationTypePerFile: modificationTypePerFile
    }
  }

  private mapNpmCommand(command: string, packagesDevExact: string[],
    packagesDepExact: string[], packagesDepUn: string[], packagesDevUn: string[]): void {
    const npmReduceRegex: RegExp = /npm\s+(i|un)\s+([\w\d\@\/\.-]+)\s+(-DE|-SE|-S|-D)?/gm;
    const npmReduceMatch: RegExpExecArray | null = npmReduceRegex.exec(command);
    const packageNameGroupId: number = 2;
    const installCommandGroupId: number = 1;
    const dependencyCategoryGroupId: number = 3;
    if (!npmReduceMatch) {
      return;
    }

    if (npmReduceMatch[installCommandGroupId] === 'i') {
      if (npmReduceMatch[dependencyCategoryGroupId] === '-SE') {
        packagesDepExact.push(npmReduceMatch[packageNameGroupId]);
      }
      else if (npmReduceMatch[dependencyCategoryGroupId] === '-DE') {
        packagesDevExact.push(npmReduceMatch[packageNameGroupId]);
      }
    }
    else {
      if (npmReduceMatch[dependencyCategoryGroupId] === '-S') {
        packagesDepUn.push(npmReduceMatch[packageNameGroupId]);
      }
      else if (npmReduceMatch[dependencyCategoryGroupId] === '-D') {
        packagesDevUn.push(npmReduceMatch[packageNameGroupId]);
      }
    }
  }

  private reduceNpmCommand(packagesDepExact: string[], packagesDevExact: string[],
    packagesDepUn: string[], packagesDevUn: string[]): string[] {
    const commandsToExecute: string[] = [];

    if (packagesDepExact.length > 0) {
      commandsToExecute.push(`npm i ${packagesDepExact.join(' ')} -SE`);
    }

    if (packagesDevExact.length > 0) {
      commandsToExecute.push(`npm i ${packagesDevExact.join(' ')} -DE`);
    }

    if (packagesDepUn.length > 0) {
      commandsToExecute.push(`npm un ${packagesDepUn.join(' ')} -S`);
    }

    if (packagesDevUn.length > 0) {
      commandsToExecute.push(`npm un ${packagesDevUn.join(' ')} -D`);
    }

    return commandsToExecute;
  }

  private getProjectRoot(folderPath: string): string | null {
    const packageJsonPath: string = path.resolve(folderPath, 'package.json');
    if (fs.existsSync(packageJsonPath)) {
      return folderPath;
    }
    else {
      const parentPath: string = path.resolve(folderPath, `..${path.sep}`);
      if (parentPath !== folderPath) {
        return this.getProjectRoot(parentPath);
      }
      else {
        return null;
      }
    }
  }

  private getProjectVersion(): string | undefined {
    const yoRcPath: string = path.resolve(this.projectRootPath as string, '.yo-rc.json');
    if (fs.existsSync(yoRcPath)) {
      try {
        const yoRc: any = JSON.parse(fs.readFileSync(yoRcPath, 'utf-8'));
        if (yoRc && yoRc['@microsoft/generator-sharepoint'] &&
          yoRc['@microsoft/generator-sharepoint'].version) {
          return yoRc['@microsoft/generator-sharepoint'].version;
        }
      }
      catch { }
    }

    const packageJsonPath: string = path.resolve(this.projectRootPath as string, 'package.json');
    try {
      const packageJson: any = JSON.parse(fs.readFileSync(packageJsonPath, 'utf-8'));
      if (packageJson &&
        packageJson.dependencies &&
        packageJson.dependencies['@microsoft/sp-core-library']) {
        const coreLibVersion: string = packageJson.dependencies['@microsoft/sp-core-library'];
        return coreLibVersion.replace(/[^0-9\.]/g, '');
      }
    }
    catch { }

    return undefined;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-v, --toVersion [toVersion]',
        description: 'The version of SharePoint Framework to which upgrade the project'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    parentOptions.forEach(o => {
      if (o.option.indexOf('--output') > -1) {
        o.description = 'Output type. json|text|md. Default text';
        o.autocomplete = ['json', 'text', 'md'];
      }
    })
    return options.concat(parentOptions);
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
    (v1.6.0).

    This command doesn't change your project files. Instead, it gives you
    a report with all steps necessary to upgrade your project to the specified
    version of the SharePoint Framework. Changing project files is error-prone,
    especially when it comes to updating your solution's code. This is why at
    this moment, this command produces a report that you can use yourself to
    perform the necessary updates and verify that everything is working as
    expected.

    Using this command you can upgrade SharePoint Framework projects built using
    versions: 1.0.0, 1.0.1, 1.0.2, 1.1.0, 1.1.1, 1.1.3, 1.2.0, 1.3.0, 1.3.1,
    1.3.2, 1.3.4, 1.4.0, 1.4.1, 1.5.0, 1.5.1 and 1.6.0.

  Examples:
  
    Get instructions to upgrade the current SharePoint Framework project to
    SharePoint Framework version 1.5.0 and save the findings in a Markdown file
      ${chalk.grey(config.delimiter)} ${this.name} --toVersion 1.5.0 --output md > upgrade-report.md

    Get instructions to Upgrade the current SharePoint Framework project to
    SharePoint Framework version 1.5.0 and show the summary of the findings
    in the shell
      ${chalk.grey(config.delimiter)} ${this.name} --toVersion 1.5.0

    Get instructions to upgrade the current SharePoint Framework project to the
    latest SharePoint Framework version supported by the Office 365 CLI
      ${chalk.grey(config.delimiter)} ${this.name}
`);
  }
}

module.exports = new SpfxProjectUpgradeCommand();