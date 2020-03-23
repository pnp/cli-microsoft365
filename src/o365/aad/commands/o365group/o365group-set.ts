import commands from '../../commands';
import * as fs from 'fs';
import * as path from 'path';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import GraphCommand from '../../../base/GraphCommand';
import { GroupUpdateService, Options } from '../../services/GroupUpdateService';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

class AadO365GroupSetCommand extends GraphCommand {

  public get name(): string {
    return commands.O365GROUP_SET;
  }

  public get description(): string {
    return 'Updates Office 365 Group properties';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    GroupUpdateService.UpdateGroup(cmd, this.resource, args.options, this.verbose, this.debug, cb, this.handleRejectedODataJsonPromise);
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>',
        description: 'The ID of the Office 365 Group to update'
      },
      {
        option: '-n, --displayName [displayName]',
        description: 'Display name for the Office 365 Group'
      },
      {
        option: '-d, --description [description]',
        description: 'Description for the Office 365 Group'
      },
      {
        option: '--owners [owners]',
        description: 'Comma-separated list of Office 365 Group owners to add'
      },
      {
        option: '--members [members]',
        description: 'Comma-separated list of Office 365 Group members to add'
      },
      {
        option: '--mailNickName [mailNickName]',
        description: 'The mail alias for the Microsoft Teams team'
      },
      {
        option: '--classification [classification]',
        description: 'The classification for the Microsoft Teams team'
      },
      {
        option: '--isPrivate [isPrivate]',
        description: 'Set to true if the Office 365 Group should be private and to false if it should be public (default)'
      },
      {
        option: '-l, --logoPath [logoPath]',
        description: 'Local path to the image file to use as group logo'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.displayName &&
        !args.options.description &&
        !args.options.members &&
        !args.options.owners &&
        typeof args.options.isPrivate === 'undefined' &&
        !args.options.logoPath) {
        return 'Specify at least one property to update';
      }

      if (!args.options.id) {
        return 'Required option id missing';
      }

      if (!Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      if (args.options.owners) {
        let owners: string[] = args.options.owners.split(',').map(o => o.trim());
        for (let i = 0; i < owners.length; i++) {
          if (owners[i].indexOf('@') < 0) {
            return `${owners[i]} is not a valid userPrincipalName`;
          }
        }
      }

      if (args.options.members) {
        let members: string[] = args.options.members.split(',').map(m => m.trim());
        for (let i = 0; i < members.length; i++) {
          if (members[i].indexOf('@') < 0) {
            return `${members[i]} is not a valid userPrincipalName`;
          }
        }
      }

      if (typeof args.options.isPrivate !== 'undefined' &&
        args.options.isPrivate !== 'true' &&
        args.options.isPrivate !== 'false') {
        return `${args.options.isPrivate} is not a valid boolean value`;
      }

      if (args.options.logoPath) {
        const fullPath: string = path.resolve(args.options.logoPath);

        if (!fs.existsSync(fullPath)) {
          return `File '${fullPath}' not found`;
        }

        if (fs.lstatSync(fullPath).isDirectory()) {
          return `Path '${fullPath}' points to a directory`;
        }
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    When updating group's owners and members, the command will add newly
    specified users to the previously set owners and members. The previously
    set users will not be replaced.

    When specifying the path to the logo image you can use both relative and
    absolute paths. Note, that ~ in the path, will not be resolved and will most
    likely result in an error.

  Examples:

    Update Office 365 Group display name
      ${this.name} --id 28beab62-7540-4db1-a23f-29a6018a3848 --displayName Finance

    Change Office 365 Group visibility to public
      ${this.name} --id 28beab62-7540-4db1-a23f-29a6018a3848 --isPrivate false

    Add new Office 365 Group owners
      ${this.name} --id 28beab62-7540-4db1-a23f-29a6018a3848 --owners "DebraB@contoso.onmicrosoft.com,DiegoS@contoso.onmicrosoft.com"

    Add new Office 365 Group members
      ${this.name} --id 28beab62-7540-4db1-a23f-29a6018a3848 --members "DebraB@contoso.onmicrosoft.com,DiegoS@contoso.onmicrosoft.com"

    Set Office 365 Group classification as MBI
      ${this.name} --id '28beab62-7540-4db1-a23f-29a6018a3848' --classification MBI

    Update Office 365 Group logo
      ${this.name} --id 28beab62-7540-4db1-a23f-29a6018a3848 --logoPath images/logo.png
`);
  }
}

module.exports = new AadO365GroupSetCommand();
