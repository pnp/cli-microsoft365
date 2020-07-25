import { CommandOption, CommandValidate } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import YammerCommand from '../../../base/YammerCommand';
import commands from '../../commands';
import { CommandInstance } from '../../../../cli';

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
      if (args.options.id && typeof args.options.id !== 'number') {
        return `${args.options.id} is not a number`;
      }

      if (args.options.userId && typeof args.options.userId !== 'number') {
        return `${args.options.userId} is not a number`;
      }

      return true;
    };
  }
}

module.exports = new YammerGroupUserAddCommand();