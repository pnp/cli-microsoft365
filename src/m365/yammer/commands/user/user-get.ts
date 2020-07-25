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
  userId?: number;
  email?: string;
}

class YammerUserGetCommand extends YammerCommand {
  public get name(): string {
    return `${commands.YAMMER_USER_GET}`;
  }

  public get description(): string {
    return 'Retrieves the current user or searches for a user by ID or e-mail';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.userId = args.options.userId !== undefined;
    telemetryProps.email = args.options.email !== undefined;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    let endPoint = `${this.resource}/v1/users/current.json`;

    if (args.options.userId) {
      endPoint = `${this.resource}/v1/users/${encodeURIComponent(args.options.userId)}.json`;
    } else if (args.options.email) {
      endPoint = `${this.resource}/v1/users/by_email.json?email=${encodeURIComponent(args.options.email)}`;
    }

    const requestOptions: any = {
      url: endPoint,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json;odata=nometadata'
      },
      json: true
    };

    request
      .get(requestOptions)
      .then((res: any): void => {
        if (args.options.output === 'json') {
          cmd.log(res);
        }
        else {
          if (res instanceof Array) {
            cmd.log((res as any[]).map((n: any) => {
              const item: any = {
                id: n.id,
                full_name: n.full_name,
                email: n.email,
                job_title: n.job_title,
                state: n.state,
                url: n.url
              };
              return item;
            }));
          } else {
            cmd.log({ id: res.id, full_name: res.full_name, email: res.email, job_title: res.job_title, state: res.state, url: res.url });
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --userId [userId]',
        description: 'Retrieve a user by ID'
      },
      {
        option: '--email [email]',
        description: 'Retrieve a user by e-mail'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (args.options.userId !== undefined && args.options.email !== undefined) {
        return `You are only allowed to search by ID or e-mail but not both`;
      }

      return true;
    };
  }
}

module.exports = new YammerUserGetCommand();