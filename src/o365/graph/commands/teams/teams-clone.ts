import * as request from 'request-promise-native';
import auth from '../../GraphAuth';
import Utils from '../../../../Utils';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import { CommandOption, CommandValidate } from '../../../../Command';
import GraphCommand from '../../GraphCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
  displayName: string;
  mailNickname: string;
  partsToClone: string;
  description?: string;
  classification?: string;
  visibility?: string;
}

class GraphTeamsCloneCommand extends GraphCommand {

  public get name(): string {
    return `${commands.TEAMS_CLONE}`;
  }

  public get description(): string {
    return 'Creates a copy of a Microsoft Teams team';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.description = typeof args.options.description !== 'undefined';
    telemetryProps.classification = typeof args.options.classification !== 'undefined';
    telemetryProps.visibility = typeof args.options.visibility !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((): request.RequestPromise => {

        let body: any = {};

        body.displayName = args.options.displayName;
        body.mailNickname = args.options.mailNickname;
        body.partsToClone = args.options.partsToClone;

        if (args.options.description) {
          body.description = args.options.description;
        }

        if (args.options.classification) {
          body.classification = args.options.classification;
        }

        if (args.options.visibility) {
          body.visibility = args.options.visibility;
        }

        const requestOptions: any = {
          url: `${auth.service.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}/clone`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${auth.service.accessToken}`,
            'content-type': 'application/json;odata=nometadata',
            'accept': 'application/json;odata.metadata=none'
          }),
          json: true,
          body: body
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((): void => {
        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (res: any): void => this.handleRejectedODataJsonPromise(res, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId <teamId>',
        description: 'The ID of the Microsoft Teams team to clone'
      },
      {
        option: '-n, --displayName',
        description: 'The display name for the new Microsoft Teams Team'
      },
      {
        option: '-m, --mailNickname <mailNickname>',
        description: 'The mail alias for the new Microsoft Teams Team'
      },
      {
        option: '-p --partsToClone <partsToClone>',
        description: 'A comma-seperated list of the parts to clone. Allowed values are apps|channels|members|settings|tabs'
      },
      {
        option: '-d --description <description>',
        description: 'The description for the new Microsoft Teams Team. Will be left blank if not specified'
      },
      {
        option: '-c --classification <classification>',
        description: 'The classification for the new Microsoft Teams Team. If not specified, will be copied from the original Microsoft Teams Team'
      },
      {
        option: '-v --visibility <visibility>',
        description: 'Specify the visibility of the new Microsoft Teams Team. Allowed values are Private|Public. If not specified, the visibility will be copied from the original Microsoft Teams Team'
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

      if (!Utils.isValidGuid(args.options.teamId)) {
        return `${args.options.teamId} is not a valid GUID`;
      }

      if (!args.options.displayName) {
        return 'Required option displayName missing';
      }

      if (!args.options.mailNickname) {
        return 'Required option mailNickname missing';
      }

      if (!args.options.partsToClone) {
        return 'Required option partsToClone missing';
      }

      if (args.options.partsToClone) {
        let partsToClone: string[] = args.options.partsToClone.split(',').map(p => p.trim());

        for (let partToClone of partsToClone) {

          if (!partToClone) {
            return `partsToClone can not have empty/blank value. Allowed values are apps|channels|members|settings|tabs`;
          }

          let part: string = partToClone.toLowerCase();

          if (part !== 'apps' &&
            part !== 'channels' &&
            part !== 'members' &&
            part !== 'settings' &&
            part !== 'tabs') {
            return `${part} is not a valid partsToClone. Allowed values are apps|channels|members|settings|tabs`;
          }
        }
      }

      if (args.options.visibility) {
        const visibility: string = args.options.visibility.trim().toLowerCase();

        if (visibility !== 'private' &&
          visibility !== 'public') {
          return `${args.options.visibility} is not a valid visibility type. Allowed values are Private|Public`;
        }
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} Before using this command, log in to the Microsoft Graph,
    using the ${chalk.blue(commands.LOGIN)} command.
          
  Remarks:
          
    To clone a Microsoft Teams team, you have to first log in to the Microsoft
    Graph using the ${chalk.blue(commands.LOGIN)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN}`)}.

    Using this command, global admins and Microsoft Teams service admins can
    access teams that they are not a member of.

    When tabs are cloned, they are put into an unconfigured state and they are 
    displayed on the tab bar in Microsoft Teams, and the first time you open them, 
    you'll go through the configuration screen. (If the person opening the tab does 
    not have permission to configure apps, they will see a message explaining that 
    the tab hasn't been configured.)

  Examples:
    
    Creates a copy of a Microsoft Teams team
      ${chalk.grey(config.delimiter)} ${commands.TEAMS_CLONE} --teamId 6f6fd3f7-9ba5-4488-bbe6-a789004d0d55 --displayName "Library Assist" --mailNickname "libassist" --partsToClone "apps,tabs,settings,channels,members" 
    
    Creates a copy of a Microsoft Teams team
      ${chalk.grey(config.delimiter)} ${commands.TEAMS_CLONE} --teamId 6f6fd3f7-9ba5-4488-bbe6-a789004d0d55 --displayName "Library Assist" --mailNickname "libassist" --partsToClone "apps,tabs,settings,channels,members" --description "Self help community for library" --classification "Library" --visibility "public"
      
    `);
  }
}

module.exports = new GraphTeamsCloneCommand();