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

        const requestOptions: any = {
          url: `${auth.service.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}/clone`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${auth.service.accessToken}`,
            "content-type": "application/zip",
            accept: 'application/json;odata.metadata=none'
          }),
          json: true,
          body: {
            displayName : args.options.displayName,
            mailNickname : args.options.mailNickname,
            partsToClone : args.options.partsToClone,
            description : args.options.description || undefined,
            classification : args.options.classification || undefined,
            visibility : args.options.visibility || undefined,
          }
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((res: any): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

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
        option: '-n, --displayName <displayName>',
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
        option: '-d --description [description]',
        description: 'The description for the new Microsoft Teams Team. Will be left blank if not specified'
      },
      {
        option: '-c --classification [classification]',
        description: 'The classification for the new Microsoft Teams Team. If not specified, will be copied from the original Microsoft Teams Team'
      },
      {
        option: '-v --visibility [visibility]',
        description: 'Specify the visibility of the new Microsoft Teams Team. Allowed values are Private|Public. If not specified, the visibility will be copied from the original Microsoft Teams Team',
        autocomplete: ['Private', 'Public']
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
    
    Creates a copy of a Microsoft Teams team with mandatory parameters
      ${chalk.grey(config.delimiter)} ${commands.TEAMS_CLONE} --teamId 15d7a78e-fd77-4599-97a5-dbb6372846c5 --displayName "Library Assist" --mailNickname "libassist" --partsToClone "apps,tabs,settings,channels,members" 
    
    Creates a copy of a Microsoft Teams team with mandatory and optional parameters
      ${chalk.grey(config.delimiter)} ${commands.TEAMS_CLONE} --teamId 15d7a78e-fd77-4599-97a5-dbb6372846c5 --displayName "Library Assist" --mailNickname "libassist" --partsToClone "apps,tabs,settings,channels,members" --description "Self help community for library" --classification "Library" --visibility "public"
      
    `);
  }
}

module.exports = new GraphTeamsCloneCommand();

