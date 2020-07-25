import commands from '../../commands';
import request from '../../../../request';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';
import * as fs from 'fs';
import * as path from 'path';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

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
        'template@odata.bind': `https://graph.microsoft.com/beta/teamsTemplates('standard')`,
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
      .then((res: any): Promise<TeamsAsyncOperation> => {
        const requestOptions: any = {
          url: `${this.resource}/beta${res.headers.location}`,
          headers: {
            accept: 'application/json;odata.metadata=minimal'
          },
          json: true
        };

        return new Promise((resolve, reject) => {
          request.get<TeamsAsyncOperation>(requestOptions)
            .then((teamsAsyncOperation: TeamsAsyncOperation) => {
              if (!args.options.wait) {
                resolve(teamsAsyncOperation);
              } else {
                this.timeout = setTimeout(() => {
                  this.waitUntilFinished(requestOptions, resolve, reject, cmd, this.dots, this.timeout)
                }, this.pollingInterval);
              }
            });
        });
      }).then((teamsAsyncOperation: TeamsAsyncOperation) => {
        if (teamsAsyncOperation.status !== TeamsAsyncOperationStatus.Succeeded) {
          return Promise.resolve(teamsAsyncOperation);
        }

        return request.get({
          url: `${this.resource}/v1.0/groups/${teamsAsyncOperation.targetResourceId}`,
          headers: {
            accept: 'application/json;odata.metadata=minimal'
          },
          json: true
        });
      })
      .then((output: any) => {
        cmd.log(output);
        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }
        cb();
      }, (err: any): void => {
        this.handleRejectedODataJsonPromise(err, cmd, cb)
      });
  }

  private waitUntilFinished(requestOptions: any, resolve: (teamsAsyncOperation: TeamsAsyncOperation) => void, reject: (error: any) => void, cmd: CommandInstance, dots?: string, timeout?: NodeJS.Timer): void {
    if (!this.debug && this.verbose) {
      dots += '.';
      process.stdout.write(`\r${dots}`);
    }

    request
      .get<TeamsAsyncOperation>(requestOptions)
      .then((teamsAsyncOperation: TeamsAsyncOperation): void => {
        if (teamsAsyncOperation.status === TeamsAsyncOperationStatus.Succeeded) {
          if (this.verbose) {
            process.stdout.write('\n');
          }
          resolve(teamsAsyncOperation);
          return;
        }
        if (teamsAsyncOperation.error) {
          reject(teamsAsyncOperation.error);
          return;
        }
        timeout = setTimeout(() => {
          this.waitUntilFinished(requestOptions, resolve, reject, cmd, dots)
        }, this.pollingInterval);
      }).catch(err => reject(err));
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
}

module.exports = new TeamsTeamAddCommand();