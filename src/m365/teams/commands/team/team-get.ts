import { Team } from '@microsoft/microsoft-graph-types';
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
  id?: string;
  name?: string;
}

class TeamsTeamGetCommand extends GraphCommand {
  public get name(): string {
    return commands.TEAM_GET;
  }

  public get description(): string {
    return 'Retrieve information about the specified Microsoft Team';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = typeof args.options.id !== 'undefined';
    telemetryProps.name = typeof args.options.name !== 'undefined';
    return telemetryProps;
  }

  private getTeamId(args: CommandArgs): Promise<string> {
    if (args.options.id) {
      return Promise.resolve(args.options.id);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/groups?$filter=displayName eq '${encodeURIComponent(args.options.name as string)}'&$select=id,resourceProvisioningOptions`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: [{ id: string, resourceProvisioningOptions: string[] }] }>(requestOptions)
      .then(response => {
        const groupItem: { id: string, resourceProvisioningOptions: string[] } | undefined = response.value[0];

        if (!groupItem) {
          return Promise.reject(`The specified team does not exist in the Microsoft Teams`);
        }

        if (groupItem.resourceProvisioningOptions.indexOf('Team') === -1) {
          return Promise.reject(`The specified team does not exist in the Microsoft Teams`);
        }

        if (response.value.length >= 2) {
          return Promise.reject(`Multiple Microsoft Teams teams with name ${args.options.name} found: ${response.value.map(x => x.id)}`);
        }

        return Promise.resolve(groupItem.id);
      });
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getTeamId(args)
      .then((teamId: string): Promise<Team> => {
        const requestOptions: any = {
          url: `${this.resource}/v1.0/teams/${encodeURIComponent(teamId)}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };
        return request.get<Team>(requestOptions);
      })
      .then((res: Team): void => {
        logger.log(res);
        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --name [name]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.id && args.options.name) {
      return 'Specify either teamId or teamName, but not both.';
    }

    if (!args.options.id && !args.options.name) {
      return 'Specify teamId or teamName, one is required';
    }

    if (args.options.id && !Utils.isValidGuid(args.options.id as string)) {
      return `${args.options.id} is not a valid GUID`;
    }
    
    return true;
  }
}

module.exports = new TeamsTeamGetCommand();