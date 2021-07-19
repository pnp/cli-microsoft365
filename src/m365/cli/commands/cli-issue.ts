import { ChildProcess } from 'child_process';
import * as open from 'open';
import { Logger } from '../../../cli';
import {
  CommandOption
} from '../../../Command';
import GlobalOptions from '../../../GlobalOptions';
import AnonymousCommand from '../../base/AnonymousCommand';
import commands from '../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  type: string;
}

class CliIssueCommand extends AnonymousCommand {
  constructor(private open: any) {
    super();
  }

  public get name(): string {
    return commands.ISSUE;
  }

  public get description(): string {
    return 'Returns, or opens a URL that takes the user to the right place in the CLI GitHub repo to create a new issue reporting bug, feedback, ideas, etc.';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.type = args.options.type;
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    let issueLink: string = '';

    switch (encodeURIComponent(args.options.type)) {
      case 'bug':
        issueLink = 'https://aka.ms/cli-m365/bug';
        break;
      case 'command':
        issueLink = 'https://aka.ms/cli-m365/new-command';
        break;
      case 'sample':
        issueLink = 'https://aka.ms/cli-m365/new-sample-script';
        break;
    }

    this.openBrowser(issueLink).then((): void => {
      logger.log(issueLink);
      cb();
    });
  }

  private async openBrowser(issueLink: string): Promise<ChildProcess> {
    return this.open(issueLink, { wait: false });
  }

  private static issueType: string[] = [
    'bug',
    'command',
    'sample'
  ];

  public validate(args: CommandArgs): boolean | string {
    if (CliIssueCommand.issueType.indexOf(args.options.type) < 0) {
      return `${args.options.type} is not a valid Issue type. Allowed values are ${CliIssueCommand.issueType.join(', ')}`;
    }

    return true;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-t, --type <type>',
        autocomplete: CliIssueCommand.issueType
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new CliIssueCommand(open);
