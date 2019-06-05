import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import GraphCommand from "../../../base/GraphCommand";
import request from '../../../../request';
import Utils from '../../../../Utils';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  description?: string;
  teamId: string;
  name: string;
}

class TeamsChannelAddCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TEAMS_CHANNEL_ADD}`;
  }

  public get description(): string {
    return 'Adds a channel to the specified Microsoft Teams team';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.description = typeof args.options.description !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/teams/${args.options.teamId}/channels`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json;odata=nometadata'
      },
      body: {
        displayName: args.options.name,
        description: args.options.description || null
      },
      json: true
    };

    request
      .post(requestOptions)
      .then((res: any): void => {
        cmd.log(res);

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId <teamId>',
        description: 'The ID of the team to add the channel to'
      },
      {
        option: '-n, --name <name>',
        description: 'The name of the channel to add'
      },
      {
        option: '-d, --description [description]',
        description: 'The description of the channel to add'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.teamId) {
        return 'Required parameter teamId missing';
      }

      if (!Utils.isValidGuid(args.options.teamId as string)) {
        return `${args.options.teamId} is not a valid GUID`;
      }

      if (!args.options.name) {
        return 'Required parameter name missing';
      }

      return true;
    };
  }


  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    You can only add a channel to the Microsoft Teams team you are a member of.

  Examples:
  
    Add channel to the specified Microsoft Teams team
      ${this.name} --teamId 6703ac8a-c49b-4fd4-8223-28f0ac3a6402 --name office365cli --description development
`   );
  }
}

module.exports = new TeamsChannelAddCommand();