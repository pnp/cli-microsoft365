import { CommandOption, CommandValidate } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import YammerCommand from '../../../base/YammerCommand';
import commands from '../../commands';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: number;
  userId?: number;
  email?: string;
}

class YammerGroupUserAddCommand extends YammerCommand {
  public get name(): string {
    return `${commands.YAMMER_GROUP_USER_ADD}`;
  }

  public get description(): string {
    return 'Adds a user to a Yammer Group';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.userId = typeof args.options.userId !== 'undefined';
    telemetryProps.email = typeof args.options.email !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const requestOptions: any = {
      url: `${this.resource}/v1/group_memberships.json`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json;odata=nometadata'
      },
      json: true,
      body: {
        group_id: args.options.id,
        user_id: args.options.userId,
        email: args.options.email
      }
    };

    request
      .post(requestOptions)
      .then((res: any): void => {
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  };

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--id <id>',
        description: 'The ID of the group to add the user to'
      },
      {
        option: '--userId [userId]',
        description: 'ID of the user to add to the group. If not specified, adds the current user'
      },
      {
        option: '--email [email]',
        description: 'E-mail of the user to add to the group'
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

    If the specified user is not a member of the network, the command will
    return an HTTP 400 error message.

  Examples:
    
    Adds the current user to the group with the ID ${chalk.grey('5611239081')}
      ${this.name} --id 5611239081
    
    Adds the user with ID ${chalk.grey('66622349')} to the group with ID ${chalk.grey('5611239081')}
      ${this.name} --id 5611239081 --userId 66622349

    Adds the user with e-mail ${chalk.grey('suzy@contoso.com')} to the group with ID
    ${chalk.grey('5611239081')}
      ${this.name} --id 5611239081 --email suzy@contoso.com
  `);
  }
}

module.exports = new YammerGroupUserAddCommand();