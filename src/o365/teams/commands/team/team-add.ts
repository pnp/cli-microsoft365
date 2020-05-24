import commands from '../../commands';
import aadcommands from '../../../aad/commands';
import request from '../../../../request';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate, CommandCancel
} from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';
import * as fs from 'fs';
import * as path from 'path';
const vorpal: Vorpal = require('../../../../vorpal-init');

enum TeamsAsyncOperationStatus {
  Invalid = "invalid",
  NotStarted = "notStarted",
  InProgress = "inProgress",
  Succeeded = "succeeded",
  Failed = "failed"
}

interface TeamsAsyncOperation {
  id: string;
  operationType: string;
  createdDateTime: Date;
  status: TeamsAsyncOperationStatus;
  lastActionDateTime: Date;
  attemptsCount: number;
  targetResourceId: string;
  targetResourceLocation: string;
  error?: any;
}

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  description?: string;
  name?: string;
  templatePath?: string;
  wait: boolean;
}

class TeamsTeamAddCommand extends GraphCommand {
  private dots?: string;
  private timeout?: NodeJS.Timer;
  private pollingInterval: number = 30000;

  public get name(): string {
    return `${commands.TEAMS_TEAM_ADD}`;
  }

  public get description(): string {
    return 'Adds a new Microsoft Teams team';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.name = typeof args.options.name !== 'undefined';
    telemetryProps.description = typeof args.options.description !== 'undefined';
    telemetryProps.templatePath = typeof args.options.templatePath !== 'undefined';
    telemetryProps.wait = args.options.wait;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    this.dots = '';

    let requestBody: any;
    if (args.options.templatePath) {
      const fullPath: string = path.resolve(args.options.templatePath);

      if (this.verbose) {
        cmd.log(`Using template '${fullPath}'...`);
      };
      requestBody = JSON.parse(fs.readFileSync(fullPath, 'utf-8'))

      if (args.options.name) {
        if (this.verbose) {
          cmd.log(`Using '${args.options.name}' as name...`);
        };
        requestBody.displayName = args.options.name;
      }

      if (args.options.description) {
        if (this.verbose) {
          cmd.log(`Using '${args.options.description}' as description...`);
        };
        requestBody.description = args.options.description;
      }

    }
    else {
      requestBody = {
        'template@odata.bind': 'https://graph.microsoft.com/beta/teamsTemplates/standard',
        displayName: args.options.name,
        description: args.options.description
      }
    }

    const requestOptions = {
      url: `${this.resource}/beta/teams`,
      resolveWithFullResponse: true,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json;odata.metadata=none'
      },
      body: requestBody,
      json: true
    };

    request
      .post(requestOptions)
      .then((response: any) => {
        return new Promise((resolve, reject) => {
          if (!args.options.wait) {
            resolve();
          }
          else {
            this.timeout = setTimeout(() => {
              this.waitUntilFinished(response.headers.location, resolve, reject, cmd, this.dots, this.timeout)
            }, this.pollingInterval);
          };
        })
      }).then(() => {
        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }
        cb();
      }, (err: any): void => {
        this.handleRejectedODataJsonPromise(err, cmd, cb)
      });
  }

  private waitUntilFinished(operationUrl: string, resolve: () => void, reject: (error: any) => void, cmd: CommandInstance, dots?: string, timeout?: NodeJS.Timer): void {
    if (!this.debug && this.verbose) {
      dots += '.';
      process.stdout.write(`\r${dots}`);
    }

    const requestOptions: any = {
      url: `${this.resource}/beta${operationUrl}`,
      headers: {
        accept: 'application/json;odata.metadata=minimal'
      },
      json: true
    };

    request
      .get<TeamsAsyncOperation>(requestOptions)
      .then((response: TeamsAsyncOperation): void => {
        if (response.status === TeamsAsyncOperationStatus.Succeeded) {
          if (this.verbose) {
            process.stdout.write('\n');
          }
          resolve();
          return;
        }
        if (response.status === TeamsAsyncOperationStatus.Invalid || response.status === TeamsAsyncOperationStatus.Failed) {
          reject(response.status);
          return;
        }
        timeout = setTimeout(() => {
          this.waitUntilFinished(operationUrl, resolve, reject, cmd, dots)
        }, this.pollingInterval);
      }).catch(err => reject(err));
  }

  public cancel(): CommandCancel {
    return (): void => {
      if (this.timeout) {
        clearTimeout(this.timeout);
      }
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --name [name]',
        description: 'Display name for the Microsoft Teams team. Required if templatePath not supplied'
      },
      {
        option: '-d, --description [description]',
        description: 'Description for the Microsoft Teams team. Required if templatePath not supplied'
      },
      {
        option: '--templatePath [templatePath]',
        description: 'Local path to the file containing the template. If name or description are supplied, these take precedence over the template values'
      },
      {
        option: '--wait',
        description: 'Wait for the team to be provisioned before completing the command'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.templatePath) {
        if (!args.options.name) {
          return `Required parameter name missing`
        }

        if (!args.options.description) {
          return `Required parameter description missing`
        }
      }

      if (args.options.templatePath && !fs.existsSync(args.options.templatePath)) {
        return 'Specified path of the template does not exist';
      };

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    ${chalk.yellow('Attention:')} This command is based on an API that is currently in preview
    and is subject to change once the API reached general availability.

    If you want to add a Team to an existing Office 365 Group use the
    ${chalk.blue(aadcommands.O365GROUP_TEAMIFY)} command instead.

  Examples:
  
    Add a new Microsoft Teams team 
      ${this.name} --name 'Architecture' --description 'Architecture Discussion'

    Add a new Microsoft Teams team using a template
      ${this.name} --name 'Architecture' --description 'Architecture Discussion' --templatePath 'template.json'

    Add a new Microsoft Teams team using a template and wait for the team
    to be provisioned
      ${this.name} --name 'Architecture' --description 'Architecture Discussion' --templatePath 'template.json' --wait

  More information:

    Get started with Teams templates
      https://docs.microsoft.com/en-us/MicrosoftTeams/get-started-with-teams-templates
  `);
  }
}

module.exports = new TeamsTeamAddCommand();