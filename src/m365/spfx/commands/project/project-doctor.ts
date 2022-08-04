import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { Logger } from '../../../../cli';
import { CommandError, CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { Dictionary, Hash, packageManager } from '../../../../utils';
import commands from '../../commands';
import { BaseProjectCommand } from './base-project-command';
import { rules as genericRules } from './project-doctor/generic-rules';
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
}

class SpfxProjectDoctorCommand extends BaseProjectCommand {
  private static packageManagers: string[] = ['npm', 'pnpm', 'yarn'];

  public static ERROR_NO_PROJECT_ROOT_FOLDER: number = 1;
  public static ERROR_NO_VERSION: number = 3;
  public static ERROR_UNSUPPORTED_VERSION: number = 4;

  private allFindings: Finding[] = [];
  private packageManager: string = 'npm';
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
    '1.15.2'
  ];

  public constructor() {
    super();
  }

  public get name(): string {
    return commands.PROJECT_DOCTOR;
  }

  public get description(): string {
    return 'Validates correctness of a SharePoint Framework project';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.packageManager = args.options.packageManager || 'npm';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    this.projectRootPath = this.getProjectRoot(process.cwd());
    if (this.projectRootPath === null) {
      cb(new CommandError(`Couldn't find project root folder`, SpfxProjectDoctorCommand.ERROR_NO_PROJECT_ROOT_FOLDER));
      return;
    }

    this.packageManager = args.options.packageManager || 'npm';

    if (this.verbose) {
      logger.logToStderr('Collecting project...');
    }
    const project: Project = this.getProject(this.projectRootPath);

    if (this.debug) {
      logger.logToStderr('Collected project');
      logger.logToStderr(project);
    }

    project.version = this.getProjectVersion();
    if (!project.version) {
      cb(new CommandError(`Unable to determine the version of the current SharePoint Framework project`, SpfxProjectDoctorCommand.ERROR_NO_VERSION));
      return;
    }

    if (!this.supportedVersions.includes(project.version)) {
      cb(new CommandError(`CLI for Microsoft 365 doesn't support validating projects built using SharePoint Framework v${project.version}`, SpfxProjectDoctorCommand.ERROR_UNSUPPORTED_VERSION));
      return;
    }

    if (this.verbose) {
      logger.logToStderr(`Project built using SPFx v${project.version}`);
    }

    const rules: Rule[] = [...genericRules];

    try {
      const versionRules: Rule[] = require(`./project-doctor/doctor-${project.version}`);
      rules.push(...versionRules);
    }
    catch (e: any) {
      cb(new CommandError(e.message));
      return;
    }

    rules.forEach(r => {
      r.visit(project, this.allFindings);
    });

    if (this.packageManager === 'npm') {
      const npmDedupeRule: Rule = new FN017001_MISC_npm_dedupe();
      npmDedupeRule.visit(project, this.allFindings);
    }

    // remove superseded findings
    this.allFindings
      // get findings that supersede other findings
      .filter(f => f.supersedes.length > 0)
      .forEach(f => {
        f.supersedes.forEach(s => {
          // find the superseded finding
          const i: number = this.allFindings.findIndex(f1 => f1.id === s);
          if (i > -1) {
            // ...and remove it from findings
            this.allFindings.splice(i, 1);
          }
        });
      });

    // flatten
    const findingsToReport: FindingToReport[] = ([] as FindingToReport[]).concat.apply([], this.allFindings.map(f => {
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
    });

    switch (args.options.output) {
      case 'text':
        logger.log(this.getTextReport(findingsToReport));
        break;
      case 'tour':
        this.writeReportTourFolder(this.getTourReport(findingsToReport));
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

    const tourFilePath: string = path.join(this.projectRootPath as string, '.tours', 'validation.tour');
    fs.writeFileSync(path.resolve(tourFilePath), findingsToReport, 'utf-8');
  }

  private getTextReport(findings: FindingToReport[]): string {
    if (findings.length === 0) {
      return '✅ CLI for Microsoft 365 has found no issues in your project';
    }

    const reportData: ReportData = this.getReportData(findings);
    const s: string[] = [
      'Execute in command line', os.EOL,
      '-----------------------', os.EOL,
      reportData.packageManagerCommands.join(os.EOL), os.EOL,
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
      `# Validate project ${path.basename(this.projectRootPath as string)}`, os.EOL,
      os.EOL,
      `Date: ${(new Date().toLocaleDateString())}`, os.EOL,
      os.EOL,
      '## Findings', os.EOL,
      os.EOL
    ];

    if (findingsToReport.length === 0) {
      s.push(`✅ CLI for Microsoft 365 has found no issues in your project`, os.EOL);
    }
    else {
      s.push(...[
        `Following is the list of issues found in your project. [Summary](#Summary) of the recommended fixes is included at the end of the report.`, os.EOL,
        os.EOL,
        findingsToReport.join(''),
        '## Summary', os.EOL,
        os.EOL,
        '### Execute script', os.EOL,
        os.EOL,
        '```sh', os.EOL,
        reportData.packageManagerCommands.join(os.EOL), os.EOL,
        '```', os.EOL,
        os.EOL
      ]);
    }

    return s.join('').trim();
  }

  private getTourReport(findings: FindingToReport[]): string {
    const tourFindings: FindingTour = {
      title: `Validate project ${path.basename(this.projectRootPath as string)}`,
      steps: []
    };

    findings.forEach(f => {
      const lineNumber: number = f.position && f.position.line ? f.position.line : 1;

      let resolution: string = '';
      switch (f.resolutionType) {
        case 'cmd':
          resolution = `Execute the following command:\r\n\r\n[\`${f.resolution}\`](command:codetour.sendTextToTerminal?["${f.resolution}"])`;
          break;
      }

      // Make severity uppercase for the markdown
      const sev: string = f.severity.toUpperCase();

      // Clean up the file name
      let file: string | undefined = f.file;

      if (file !== undefined) {
        // CodeTour expects the files to be relative from root (i.e.: no './')
        file = file.replace(/\.\//g, '');

        // CodeTour also expects forward slashes as directory separators
        file = file.replace(/\\/g, '/');
      }

      // Create a tour step entry
      const step: FindingTourStep = {
        file,
        title: `${sev}: ${f.title} (${f.id})`,
        description: `### ${sev}\r\n\r\n${f.description}\r\n\r\n${resolution}`,
        line: lineNumber
      };

      tourFindings.steps.push(step);
    });

    // Add the finale
    tourFindings.steps.push({
      file: ".tours/validation.tour",
      title: "RECOMMENDED: Delete tour",
      description: "### THAT'S IT!!!\r\nOnce you have tested that your project has no more issues, you can delete the `.tour` folder and its contents. Otherwise, you'll be prompted to launch this CodeTour every time you open this project."
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
      packageManager.mapPackageManagerCommand({
        command: f.resolution,
        packagesDevExact,
        packagesDepExact,
        packagesDepUn,
        packagesDevUn,
        packageMgr: this.packageManager
      });
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

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--packageManager [packageManager]',
        autocomplete: SpfxProjectDoctorCommand.packageManagers
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    parentOptions.forEach(o => {
      if (o.option.indexOf('--output') > -1) {
        o.autocomplete = ['json', 'text', 'md', 'tour'];
      }
    });
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.packageManager) {
      if (SpfxProjectDoctorCommand.packageManagers.indexOf(args.options.packageManager) < 0) {
        return `${args.options.packageManager} is not a supported package manager. Supported package managers are ${SpfxProjectDoctorCommand.packageManagers.join(', ')}`;
      }
    }

    return true;
  }
}

module.exports = new SpfxProjectDoctorCommand();
