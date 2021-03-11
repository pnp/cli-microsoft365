import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  title: string;
  ownerGroupId?: string;
  ownerGroupName?: string;
}

class GraphPlannerPlanAddCommand extends GraphCommand {
  public get name(): string {
    return commands.PLANNER_PLAN_ADD;
  }

  public get description(): string {
    return 'Adds a new Microsoft Planner plan';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.ownerGroupId = typeof args.options.ownerGroupId !== 'undefined';
    telemetryProps.ownerGroupName = typeof args.options.ownerGroupName !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const endpointPlanner: string = `${this.resource}/v1.0/planner/plans`;
    if (args.options.ownerGroupId) {
        ((): Promise<any[]> => {
          return this.getGroupsWithFilter(`?$filter=ID eq '${args.options.ownerGroupId}'`);
        })()
          .then((groups: any[]): void => {
            if (groups && groups.length > 0) {
              if (groups.length > 1) {
                logger.log(`More than one groups found with id ${args.options.ownerGroupId}`);
                cb();
              }
              else {
                const requestOptions: any = {
                  url: endpointPlanner,
                  headers: {
                    'accept': 'application/json;odata.metadata=none'
                  },
                  responseType: 'json',
                  data: {
                    "owner": groups[0].id,
                    "title": args.options.title
                  }
                };
                request
                  .post(requestOptions)
                  .then((res: any): void => {
                    logger.log(res);
                    cb();
                  }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
              }
            }
            else {
              logger.log(`Owner group not found with id ${args.options.ownerGroupId}`);
              cb();
            }
          }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
    }
    else {
      ((): Promise<any[]> => {
        const ownerGroupNameFilter: string = args.options.ownerGroupName ? `?$filter=DisplayName eq '${encodeURIComponent(args.options.ownerGroupName).replace(/'/g, `''`)}'` : '';
        return this.getGroupsWithFilter(ownerGroupNameFilter);
      })()
        .then((groups: any[]): void => {
          if (groups && groups.length > 0) {
            if (groups.length > 1) {
              logger.log(`More than one groups found with name ${args.options.ownerGroupName}`);
              cb();
            }
            else {
              const requestOptions: any = {
                url: endpointPlanner,
                headers: {
                  'accept': 'application/json;odata.metadata=none'
                },
                responseType: 'json',
                data: {
                  "owner": groups[0].id,
                  "title": args.options.title
                }
              };
              request
                .post(requestOptions)
                .then((res: any): void => {
                  logger.log(res);
                  cb();
                }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
            }
          }
          else {
            logger.log(`Owner group not found with name ${args.options.ownerGroupName}`);
            cb();
          }
        }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
    }
  }

  private getGroupsWithFilter(ownerGroupFilter: string): Promise<any[]> {
    const endpoint: string = `${this.resource}/v1.0/groups${ownerGroupFilter}`;
    const requestOptions: any = {
      url: endpoint,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return new Promise<any[]>((resolve: (groups: any[]) => void, reject: (error: any) => void): void => {
      request
        .get<{ value: any[]; }>(requestOptions)
        .then((response: { value: any[] }) => {
          resolve(response.value);
        })
        .catch((error: any) => {
          reject(error);
        });
    });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-t, --title <title>'
      },
      {
        option: "--ownerGroupId [ownerGroupId]"
      },
      {
        option: "--ownerGroupName [ownerGroupName]"
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.ownerGroupId && !args.options.ownerGroupName) {
      return 'Specify either ownerGroupId or ownerGroupName but not both';
    }

    if (args.options.ownerGroupId && args.options.ownerGroupName) {
      return 'Specify either ownerGroupId or ownerGroupName but not both';
    }

    if (args.options.ownerGroupId && !Utils.isValidGuid(args.options.ownerGroupId as string)) {
      return `${args.options.ownerGroupId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new GraphPlannerPlanAddCommand();