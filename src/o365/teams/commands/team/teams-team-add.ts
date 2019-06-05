import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import GraphCommand from '../../../base/GraphCommand';
import request from '../../../../request';
const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  description?: string;
  groupId?: string;
  name?: string;
}

class TeamsTeamAddCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TEAMS_TEAM_ADD}`;
  }

  public get description(): string {
    return 'Adds a new Microsoft Teams team';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.groupId = typeof args.options.groupId !== 'undefined';
    telemetryProps.name = typeof args.options.name !== 'undefined';
    telemetryProps.description = typeof args.options.description !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
      ((): Promise<{}> => {
        return args.options.groupId ? this.createTeamForGroup(args) :
          this.createTeam(args);
      })()
      .then((res: any): void => {
        // get the teams id from the response header.
        const teamsRspHdrRegEx: RegExpExecArray | null = /teams?\('([^']+)'\)/i.exec(res.headers.location);

        if (teamsRspHdrRegEx !== null && teamsRspHdrRegEx.length > 0) {
          cmd.log(teamsRspHdrRegEx[1]);
        }

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }
        cb();
      }, (err: any): void => {
        this.handleRejectedODataJsonPromise(err, cmd, cb)
      });
  }

  private createTeam(args: CommandArgs): Promise<{}> {
    const requestOptions: any = {
      url: `${this.resource}/beta/teams`,
      resolveWithFullResponse: true,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json;odata.metadata=none'
      },
      body: {
        'template@odata.bind': 'https://graph.microsoft.com/beta/teamsTemplates/standard',
        displayName: args.options.name,
        description: args.options.description
      },
      json: true
    };

    return request.post(requestOptions);
  }

  private createTeamForGroup(args: CommandArgs): Promise<{}> {
    const requestOptions: any = {
      url: `${this.resource}/beta/groups/${args.options.groupId}/team`,
      resolveWithFullResponse: true,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json;odata.metadata=none'
      },
      body: {},
      json: true
    };

    return request.put(requestOptions);
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --name [name]',
        description: 'Display name for the Microsoft Teams team. Required, when groupId is not specified.'
      },
      {
        option: '-d, --description [description]',
        description: 'Description for the Microsoft Teams team. Required, when groupId is not specified.'
      },
      {
        option: '-i, --groupId [groupId]',
        description: 'The ID of the Office 365 group to add a Microsoft Teams team to'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.groupId) {
        if (!args.options.name) {
          return `Required parameter name missing`
        }

        if (!args.options.description) {
          return `Required parameter description missing`
        }
      }
      else {
        if (args.options.name) {
          return `Specify either groupId or name but not both`;
        }

        if (args.options.description) {
          return `Specifying description with groupId is not supported`;
        }

        if (!Utils.isValidGuid(args.options.groupId as string)) {
          return `${args.options.groupId} is not a valid GUID`;
        }
      }

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

  Examples:
  
    Add a new Microsoft Teams team by creating a group 
      ${this.name} --name 'Architecture' --description 'Architecture Discussion'

    Add a new Microsoft Teams team to an existing Office 365 group 
      ${this.name} --groupId 6d551ed5-a606-4e7d-b5d7-36063ce562cc
  `);
  }
}

module.exports = new TeamsTeamAddCommand();