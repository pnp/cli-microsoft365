import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  groupId: string;
}

class AadO365GroupTeamifyCommand extends GraphCommand {
  public get name(): string {
    return `${commands.O365GROUP_TEAMIFY}`;
  }

  public get description(): string {
    return 'Creates a new Microsoft Teams team under existing Microsoft 365 group';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    
    const body: any = {
      "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
      "group@odata.bind": `https://graph.microsoft.com/v1.0/groups('${encodeURIComponent(args.options.groupId)}')`
    }

    const requestOptions: any = {
      url: `${this.resource}/beta/teams`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      body: body,
      json: true
    };

    request
      .post(requestOptions)
      .then((): void => {
        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --groupId <groupId>',
        description: 'The ID of the Microsoft 365 Group to connect to Microsoft Teams'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!Utils.isValidGuid(args.options.groupId)) {
        return `${args.options.groupId} is not a valid GUID`;
      }

      return true;
    };
  }
}

module.exports = new AadO365GroupTeamifyCommand();