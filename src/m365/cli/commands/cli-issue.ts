import { ChildProcess } from 'child_process';
import * as open from 'open';
import { Logger } from '../../../cli';
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
  public get name(): string {
    return commands.ISSUE;
  }

  public get description(): string {
    return 'Returns, or opens a URL that takes the user to the right place in the CLI GitHub repo to create a new issue reporting bug, feedback, ideas, etc.';
  }

  constructor(private open: any) {
    super();
  
    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }
  
  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        type: args.options.type
      });
    });
  }
  
  #initOptions(): void {
    this.options.unshift(
      {
        option: '-t, --type <type>',
        autocomplete: CliIssueCommand.issueType
      }
    );
  }
  
  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (CliIssueCommand.issueType.indexOf(args.options.type) < 0) {
          return `${args.options.type} is not a valid Issue type. Allowed values are ${CliIssueCommand.issueType.join(', ')}`;
        }
    
        return true;
      }
    );
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
}

module.exports = new CliIssueCommand(open);
