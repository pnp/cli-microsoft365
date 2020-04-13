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
}

class YammerGroupLeaveCommand extends YammerCommand {
  constructor() {
    super();
  }

  public get name(): string {
    return `${commands.YAMMER_GROUP_LEAVE}`;
  }

  public get description(): string {
    return 'Leave a Yammer group';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = args.options.id !== undefined;
    telemetryProps.userId = args.options.userId !== undefined;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
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

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--id <id>',
        description: 'The Group ID of the group to leave'
      },
      {
        option: '--userId [id]',
        description: 'Remove the user with the ID specified. Defaults to the current user'
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
    application used by the Office 365 CLI the permission to the Yammer API.
    To do this, execute the ${chalk.blue('consent --service yammer')} command.

  Examples:
    
    Remove the current user from the group with the ID ${chalk.grey('5611239081')}
      ${this.name} --id 5611239081
    
    Remove the the user with the ID ${chalk.grey('66622349')} from the group with the ID ${chalk.grey('5611239081')}
      ${this.name} --id 5611239081 --userId 66622349
  `);
  }
}

module.exports = new YammerGroupLeaveCommand();
