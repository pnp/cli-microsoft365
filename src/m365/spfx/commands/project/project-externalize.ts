import commands from '../../commands';
import Command, {
  CommandOption, CommandError, CommandAction
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import * as path from 'path';
import * as os from 'os';
import { Project, ExternalConfiguration, External } from './model';
import rules = require('./project-externalize/DefaultRules');
import { BasicDependencyRule } from './project-externalize/rules';
import { ExternalizeEntry, FileEdit } from './project-externalize/';
import { BaseProjectCommand } from './base-project-command';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: GlobalOptions;
}

class SpfxProjectExternalizeCommand extends BaseProjectCommand {
  public constructor() {
    super();
  }

  private projectVersion: string | undefined;
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
    '1.9.1'
  ];
  private allFindings: ExternalizeEntry[] = [];
  private allEditSuggestions: FileEdit[] = [];
  public static ERROR_NO_PROJECT_ROOT_FOLDER: number = 1;
  public static ERROR_NO_VERSION: number = 3;
  public static ERROR_UNSUPPORTED_VERSION: number = 2;

  public get name(): string {
    return commands.PROJECT_EXTERNALIZE;
  }

  public get description(): string {
    return 'Externalizes SharePoint Framework project dependencies';
  }

  public action(): CommandAction {
    const cmd: Command = this;
    return function (this: CommandInstance, args: CommandArgs, cb: (err?: any) => void) {
      (cmd as any).initAction(args, this);
      cmd.commandAction(this, args, cb);
    }
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    if (args.options.output !== 'json' || this.verbose) {
      cmd.log(`This command is currently in preview. Feedback welcome at https://github.com/pnp/cli-microsoft365/issues${os.EOL}`);
    }

    this.projectRootPath = this.getProjectRoot(process.cwd());
    if (this.projectRootPath === null) {
      cb(new CommandError(`Couldn't find project root folder`, SpfxProjectExternalizeCommand.ERROR_NO_PROJECT_ROOT_FOLDER));
      return;
    }

    this.projectVersion = this.getProjectVersion();
    if (!this.projectVersion) {
      cb(new CommandError(`Unable to determine the version of the current SharePoint Framework project`, SpfxProjectExternalizeCommand.ERROR_NO_VERSION));
      return;
    }

    if (this.projectVersion && this.supportedVersions.indexOf(this.projectVersion) < 0) {
      cb(new CommandError(`CLI for Microsoft 365 doesn't support externalizing dependencies of SharePoint Framework projects of version ${this.projectVersion}. Supported versions are ${this.supportedVersions.join(', ')}`, SpfxProjectExternalizeCommand.ERROR_UNSUPPORTED_VERSION));
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

    const asyncRulesResults = (rules as BasicDependencyRule[]).map(r => r.visit(project));
    Promise
      .all(asyncRulesResults)
      .then((rulesResults) => {
        this.allFindings.push(...rulesResults.map(x => x.entries).reduce((x, y) => [...x, ...y]));
        this.allEditSuggestions.push(...rulesResults.map(x => x.suggestions).reduce((x, y) => [...x, ...y]));
        //removing duplicates
        this.allFindings = this.allFindings.filter((x, i) => this.allFindings.findIndex(y => y.key === x.key) === i);
        this.writeReport(this.allFindings, this.allEditSuggestions, cmd, args.options);
        cb();
      }).catch((err) => {
        cb(new CommandError(err));
      });
  }

  private writeReport(findingsToReport: ExternalizeEntry[], editsToReport: FileEdit[], cmd: CommandInstance, options: GlobalOptions): void {
    let report;

    switch (options.output) {
      case 'json':
        report = { externalConfiguration: this.serializeJsonReport(findingsToReport), edits: editsToReport };
        break;
      case 'md':
        report = this.serializeMdReport(findingsToReport, editsToReport);
        break;
      default:
        report = this.serializeTextReport(findingsToReport, editsToReport);
        break;
    }

    cmd.log(report);
  }

  private serializeMdReport(findingsToReport: ExternalizeEntry[], editsToReport: FileEdit[]): string {
    const lines = [
      `# Externalizing dependencies of project ${path.basename(this.projectRootPath as string)}`, os.EOL,
      os.EOL,
      `Date: ${(new Date().toLocaleDateString())}`, os.EOL,
      os.EOL,
      '## Findings', os.EOL,
      os.EOL,
      '### Modify files', os.EOL,
      os.EOL,
      '#### [config.json](config/config.json)', os.EOL,
      os.EOL,
      'Replace the externals property (or add if not defined) with', os.EOL,
      os.EOL,
      '```json', os.EOL,
      JSON.stringify(this.serializeJsonReport(findingsToReport), null, 2), os.EOL,
      '```', os.EOL,
      ...this.getReportForFileEdit(this.getGroupedFileEdits(editsToReport, 'add')),
      ...this.getReportForFileEdit(this.getGroupedFileEdits(editsToReport, 'remove')),
    ];
    return lines.join('');
  }

  private getReportForFileEdit(suggestions: FileEdit[][]): string[] {
    const initialReport = suggestions
      .map(x => [
        `#### [${x[0].path}](${x[0].path})`, os.EOL,
        x[0].action, os.EOL,
        '```JavaScript', os.EOL,
        ...x.map(y => [y.targetValue, os.EOL]).reduce((y, z) => [...y, ...z]), '```', os.EOL]);

    if (initialReport.length > 0) {
      return initialReport.reduce((x, y) => [...x, ...y]);
    }
    else {
      return [];
    }
  }

  private getGroupedFileEdits(editsToReport: FileEdit[], action: "add" | "remove"): FileEdit[][] {
    const editsMatchingAction = editsToReport.filter(x => x.action === action);
    return editsMatchingAction
      .filter((x, i) => editsMatchingAction.findIndex(y => y.path === x.path) === i)
      .map(x => editsMatchingAction.filter(y => y.path === x.path));
  }
  private serializeJsonReport(findingsToReport: ExternalizeEntry[]): { externals: ExternalConfiguration } {
    const result: ExternalConfiguration = {};
    findingsToReport.forEach((f) => {
      if (!f.globalName) {
        result[f.key] = f.path;
      }
      else {
        result[f.key] = {
          path: f.path,
          globalName: f.globalName,
          globalDependencies: f.globalDependencies
        } as External;
      }
    });

    return {
      externals: result
    };
  }

  private serializeTextReport(findingsToReport: ExternalizeEntry[], editsToReport: FileEdit[]): string {
    const s: string[] = [
      'In the config/config.json file update the externals property to:', os.EOL,
      os.EOL,
      JSON.stringify({ externalConfiguration: this.serializeJsonReport(findingsToReport), edits: editsToReport }, null, 2)
    ];

    return s.join('').trim();
  }

  public options(): CommandOption[] {
    const parentOptions: CommandOption[] = super.options();
    parentOptions.forEach(o => {
      if (o.option.indexOf('--output') > -1) {
        o.description = 'Output type. json|text|md. Default text';
        o.autocomplete = ['json', 'text', 'md'];
      }
    });
    return parentOptions;
  }
}

module.exports = new SpfxProjectExternalizeCommand();
