import { AxiosRequestConfig } from 'axios';
import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { aadGroup } from '../../../../utils/aadGroup';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

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
  template?: string;
  wait: boolean;
}

class TeamsTeamAddCommand extends GraphCommand {
  private dots?: string;
  private pollingInterval: number = 30000;

  public get name(): string {
    return commands.TEAM_ADD;
  }

  public get description(): string {
    return 'Adds a new Microsoft Teams team';
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
        name: typeof args.options.name !== 'undefined',
        description: typeof args.options.description !== 'undefined',
        template: typeof args.options.template !== 'undefined',
        wait: args.options.wait
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name [name]'
      },
      {
        option: '-d, --description [description]'
      },
      {
        option: '--template [template]'
      },
      {
        option: '--wait'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!args.options.template) {
          if (!args.options.name) {
            return `Required parameter name missing`;
          }

          if (!args.options.description) {
            return `Required parameter description missing`;
          }
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this.dots = '';

    let requestBody: any;
    if (args.options.template) {
      if (this.verbose) {
        logger.logToStderr(`Using template...`);
      }
      requestBody = JSON.parse(args.options.template);

      if (args.options.name) {
        if (this.verbose) {
          logger.logToStderr(`Using '${args.options.name}' as name...`);
        }
        requestBody.displayName = args.options.name;
      }

      if (args.options.description) {
        if (this.verbose) {
          logger.logToStderr(`Using '${args.options.description}' as description...`);
        }
        requestBody.description = args.options.description;
      }
    }
    else {
      requestBody = {
        'template@odata.bind': `https://graph.microsoft.com/v1.0/teamsTemplates('standard')`,
        displayName: args.options.name,
        description: args.options.description
      };
    }

    const requestOptions: AxiosRequestConfig = {
      url: `${this.resource}/v1.0/teams`,
      headers: {
        'accept': 'application/json;odata.metadata=none'
      },
      data: requestBody,
      responseType: 'stream'
    };

    request
      .post(requestOptions)
      .then((res: any): Promise<TeamsAsyncOperation> => {
        const requestOptions: any = {
          url: `${this.resource}/v1.0${res.headers.location}`,
          headers: {
            accept: 'application/json;odata.metadata=minimal'
          },
          responseType: 'json'
        };

        return new Promise((resolve, reject) => {
          request.get<TeamsAsyncOperation>(requestOptions)
            .then((teamsAsyncOperation: TeamsAsyncOperation) => {
              if (!args.options.wait) {
                resolve(teamsAsyncOperation);
              }
              else {
                setTimeout(() => {
                  this.waitUntilFinished(requestOptions, resolve, reject, logger, this.dots);
                }, this.pollingInterval);
              }
            });
        });
      })
      .then((teamsAsyncOperation: TeamsAsyncOperation): any => {
        if (teamsAsyncOperation.status !== TeamsAsyncOperationStatus.Succeeded) {
          return Promise.resolve(teamsAsyncOperation);
        }

        return aadGroup.getGroupById(teamsAsyncOperation.targetResourceId);
      })
      .then((output: any) => {
        logger.log(output);
        cb();
      }, (err: any): void => {
        this.handleRejectedODataJsonPromise(err, logger, cb);
      });
  }

  private waitUntilFinished(requestOptions: any, resolve: (teamsAsyncOperation: TeamsAsyncOperation) => void, reject: (error: any) => void, logger: Logger, dots?: string): void {
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
        setTimeout(() => {
          this.waitUntilFinished(requestOptions, resolve, reject, logger, dots);
        }, this.pollingInterval);
      }).catch(err => reject(err));
  }
}

module.exports = new TeamsTeamAddCommand();