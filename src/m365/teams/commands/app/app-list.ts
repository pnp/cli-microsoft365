import { Group } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import { odata, validation } from '../../../../utils';
import { aadGroup } from '../../../../utils/aadGroup';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { TeamsApp } from '../../TeamsApp';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  all?: boolean;
  teamId?: string;
  teamName?: string;
}

interface ExtendedGroup extends Group {
  resourceProvisioningOptions: string[];
}

class TeamsAppListCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_LIST;
  }

  public get description(): string {
    return 'Lists apps from the Microsoft Teams app catalog or apps installed in the specified team';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'distributionMethod'];
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
        all: args.options.all || false,
        teamId: typeof args.options.teamId !== 'undefined',
        teamName: typeof args.options.teamName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-a, --all'
      },
      {
        option: '-i, --teamId [teamId]'
      },
      {
        option: '-t --teamName [teamName]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.teamId && args.options.teamName) {
          return 'Specify either teamId or teamName, but not both.';
        }

        if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
          return `${args.options.teamId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  private getTeamId(args: CommandArgs): Promise<string> {
    if (args.options.teamId) {
      return Promise.resolve(args.options.teamId);
    }

    return aadGroup
      .getGroupByDisplayName(args.options.teamName!)
      .then(group => {
        if ((group as ExtendedGroup).resourceProvisioningOptions.indexOf('Team') === -1) {
          return Promise.reject(`The specified team does not exist in the Microsoft Teams`);
        }

        return group.id!;
      });
  }

  private getEndpointUrl(args: CommandArgs): Promise<string> {
    return new Promise<string>((resolve: (endpoint: string) => void, reject: (error: string) => void): void => {
      if (args.options.teamId || args.options.teamName) {
        this
          .getTeamId(args)
          .then((teamId: string): void => {
            let endpoint: string = `${this.resource}/v1.0/teams/${encodeURIComponent(teamId)}/installedApps?$expand=teamsApp`;

            if (!args.options.all) {
              endpoint += `&$filter=teamsApp/distributionMethod eq 'organization'`;
            }

            return resolve(endpoint);
          })
          .catch((err: any) => {
            reject(err);
          });
      }
      else {
        let endpoint: string = `${this.resource}/v1.0/appCatalogs/teamsApps`;

        if (!args.options.all) {
          endpoint += `?$filter=distributionMethod eq 'organization'`;
        }

        return resolve(endpoint);
      }
    });
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    this
      .getEndpointUrl(args)
      .then(endpoint => odata.getAllItems<TeamsApp>(endpoint))
      .then((items): void => {
        if (args.options.teamId || args.options.teamName) {
          items.forEach(t => {
            t.displayName = (t as any).teamsApp.displayName;
            t.distributionMethod = (t as any).teamsApp.distributionMethod;
          });
        }

        logger.log(items);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new TeamsAppListCommand();