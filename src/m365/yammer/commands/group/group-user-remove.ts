import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import YammerCommand from "../../../base/YammerCommand";
import request from '../../../../request';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: number;
  userId?: number;
  confirm?: boolean;
}

class YammerGroupUserRemoveCommand extends YammerCommand {
  public get name(): string {
    return `${commands.YAMMER_GROUP_USER_REMOVE}`;
  }

  public get description(): string {
    return 'Removes a user from a Yammer group';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.userId = args.options.userId !== undefined;
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const executeRemoveAction: () => void = (): void => {
      let endpoint = `${this.resource}/v1/group_memberships.json`;

      const requestOptions: any = {
        url: endpoint,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json;odata=nometadata'
        },
        json: true,
        body: {
          group_id: args.options.id,
          user_id: args.options.userId
        }
      };

      request
        .delete(requestOptions)
        .then((res: any): void => {
          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
    };

    if (args.options.confirm) {
      executeRemoveAction();
    }
    else {
      let messagePrompt: string = `Are you sure you want to leave group ${args.options.id}?`;
      if (args.options.userId) {
        messagePrompt = `Are you sure you want to remove the user ${args.options.userId} from the group ${args.options.id}?`;
      }

      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: messagePrompt,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          executeRemoveAction();
        }
      });
    }
  };

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--id <id>',
        description: 'The ID of the Yammer group'
      },
      {
        option: '--userId [userId]',
        description: 'ID of the user to remove from the group. If not specified, removes the current user'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirmation before removing the user from the group'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.id) {
        return 'Required id value is missing';
      }

      if (args.options.id && typeof args.options.id !== 'number') {
        return `${args.options.id} is not a number`;
      }

      if (args.options.userId && typeof args.options.userId !== 'number') {
        return `${args.options.userId} is not a number`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:
  
    ${chalk.yellow('Attention:')} In order to use this command, you need to grant the Azure AD
    application used by the CLI for Microsoft 365 the permission to the Yammer API.
    To do this, execute the ${chalk.blue('cli consent --service yammer')} command.

  Examples:
    
    Remove the current user from the group with the ID ${chalk.grey('5611239081')}
      ${this.name} --id 5611239081
    
    Remove the user with the ID ${chalk.grey('66622349')} from the group with the ID
    ${chalk.grey('5611239081')}
      ${this.name} --id 5611239081 --userId 66622349

    Remove the user with the ID ${chalk.grey('66622349')} from the group with the ID
    ${chalk.grey('5611239081')} without asking for confirmation
      ${this.name} --id 5611239081 --userId 66622349 --confirm
  `);
  }
}

module.exports = new YammerGroupUserRemoveCommand();