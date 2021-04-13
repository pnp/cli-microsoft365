import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import GraphCommand from "../../../base/GraphCommand";
import commands from '../../commands';
import { Team } from '../../Team';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId?: string;
  teamName?: string;
  name: string;
  description?: string;
}

class TeamsChannelAddCommand extends GraphCommand {
  public get name(): string {
    return commands.CHANNEL_ADD;
  }

  public get description(): string {
    return 'Adds a channel to the specified Microsoft Teams team';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.description = typeof args.options.description !== 'undefined';
    telemetryProps.teamId = typeof args.options.teamId !== 'undefined';
    telemetryProps.teamName = typeof args.options.teamName !== 'undefined';
    return telemetryProps;
  }

  private getTeamId(args: CommandArgs): Promise<string> {
    if (args.options.teamId) {
      return Promise.resolve(args.options.teamId);
    }

    const teamRequestOptions: any = {
      url: `${this.resource}/v1.0/me/joinedTeams?$filter=displayName eq '${encodeURIComponent(args.options.teamName as string)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: Team[] }>(teamRequestOptions)
      .then(response => {
        const teamItem: Team | undefined = response.value[0];

        if (!teamItem) {
          return Promise.reject(`The specified team does not exist in the Microsoft Teams`);
        }

        if (response.value.length > 1) {
          return Promise.reject(`Multiple Microsoft Teams teams with name ${args.options.teamName} found: ${response.value.map(x => x.id)}`);
        }

        return Promise.resolve(teamItem.id);
      });
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getTeamId(args)
      .then((teamId: string): Promise<void> => {
        const requestOptions: any = {
          url: `${this.resource}/v1.0/teams/${teamId}/channels`,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'content-type': 'application/json;odata=nometadata'
          },
          data: {
            displayName: args.options.name
          },
          responseType: 'json'
        };

        return request.post(requestOptions);
      })
      .then((res: any): void => {
        logger.log(res);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId [teamId]'
      },
      {
        option: '--teamName [teamName]'
      },
      {
        option: '-n, --name <name>'
      },
      {
        option: '-d, --description [description]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.teamId && args.options.teamName) {
      return 'Specify either teamId or teamName, but not both.';
    }

    if (!args.options.teamId && !args.options.teamName) {
      return 'Specify teamId or teamName, one is required';
    }

    if (args.options.teamId && !Utils.isValidGuid(args.options.teamId)) {
      return `${args.options.teamId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new TeamsChannelAddCommand();