
import Utils from '../../../../Utils';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import { CommandOption, CommandValidate } from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';
import request from '../../../../request';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
  displayName: string;
  partsToClone: string;
  description?: string;
  classification?: string;
  visibility?: string;
}

class TeamsCloneCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TEAMS_TEAM_CLONE}`;
  }

  public get description(): string {
    return 'Creates a clone of a Microsoft Teams team';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.description = typeof args.options.description !== 'undefined';
    telemetryProps.classification = typeof args.options.classification !== 'undefined';
    telemetryProps.visibility = typeof args.options.visibility !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const body: any = {
      displayName: args.options.displayName,
      mailNickname: this.generateMailNickname(args.options.displayName),
      partsToClone: args.options.partsToClone,
    }
    if (args.options.description) {
      body.description = args.options.description;
    }
    if (args.options.classification) {
      body.classification = args.options.classification;
    }
    if (args.options.visibility) {
      body.visibility = args.options.visibility
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}/clone`,
      headers: {
        "content-type": "application/json",
        accept: 'application/json;odata.metadata=none'
      },
      json: true,
      body: body
    };

    request
      .post(requestOptions)
      .then((): void => {
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
        description: 'The ID of the Microsoft Teams team to clone'
      },
      {
        option: '-n, --displayName <displayName>',
        description: 'The display name for the new Microsoft Teams Team'
      },
      {
        option: '-p, --partsToClone <partsToClone>',
        description: 'A comma-separated list of the parts to clone. Allowed values are apps|channels|members|settings|tabs',
        autocomplete: ['apps', 'channels', 'members', 'settings', 'tabs']
      },
      {
        option: '-d, --description [description]',
        description: 'The description for the new Microsoft Teams Team'
      },
      {
        option: '-c, --classification [classification]',
        description: 'The classification for the new Microsoft Teams Team. If not specified, will be copied from the original Microsoft Teams Team'
      },
      {
        option: '-v, --visibility [visibility]',
        description: 'Specify the visibility of the new Microsoft Teams Team. Allowed values are Private|Public.',
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

      if (!args.options.partsToClone) {
        return 'Required option partsToClone missing';
      }

      const partsToClone: string[] = args.options.partsToClone.replace(/\s/g, '').split(',');
      for (let partToClone of partsToClone) {
        const part: string = partToClone.toLowerCase();
        if (part !== 'apps' &&
          part !== 'channels' &&
          part !== 'members' &&
          part !== 'settings' &&
          part !== 'tabs') {
          return `${part} is not a valid partsToClone. Allowed values are apps|channels|members|settings|tabs`;
        }
      }

      if (args.options.visibility) {
        const visibility: string = args.options.visibility.toLowerCase();

        if (visibility !== 'private' &&
          visibility !== 'public') {
          return `${args.options.visibility} is not a valid visibility type. Allowed values are Private|Public`;
        }
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:
          
    Using this command, global admins and Microsoft Teams service admins can
    access teams that they are not a member of.

    When tabs are cloned, they are put into an unconfigured state. The first
    time you open them, you'll go through the configuration screen.
    If the person opening the tab does not have permission to configure apps,
    they will see a message explaining that the tab hasn't been configured.

  Examples:
    
    Creates a clone of a Microsoft Teams team with mandatory parameters
      ${commands.TEAMS_TEAM_CLONE} --teamId 15d7a78e-fd77-4599-97a5-dbb6372846c5 --displayName "Library Assist" --partsToClone "apps,tabs,settings,channels,members" 
    
    Creates a clone of a Microsoft Teams team with mandatory and optional
    parameters
      ${commands.TEAMS_TEAM_CLONE} --teamId 15d7a78e-fd77-4599-97a5-dbb6372846c5 --displayName "Library Assist" --partsToClone "apps,tabs,settings,channels,members" --description "Self help community for library" --classification "Library" --visibility "public" 
    `);
  }

  /**
   * There is a know issue that the mailNickname is currently ignored and cannot be set by the user
   * However the mailNickname is still required by the payload so to deliver better user experience
   * the CLI generates mailNickname for the user 
   * so the user does not have to specify something that will be ignored.
   * For more see: https://docs.microsoft.com/en-us/graph/api/team-clone?view=graph-rest-1.0#request-body
   * This method has to be removed once the graph team fixes the issue and then the actual value
   * of the mailNickname would have to be specified by the CLI user.
   * @param displayName teams display name
   */
  private generateMailNickname(displayName: string): string {
    // currently the Microsoft Graph generates mailNickname in a similar fashion
    return `${displayName.replace(/[^a-zA-Z0-9]/g, "")}${Math.floor(Math.random() * 9999)}`;
  }
}

module.exports = new TeamsCloneCommand();