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
    var endPoint = `${this.resource}/v1/users/current.json`
    
    if (args.options.userId !== undefined || args.options.email !== undefined) {
      if (args.options.userId !== undefined)
        endPoint = `${this.resource}/v1/users/${args.options.userId}.json`
      else 
        endPoint = `${this.resource}/v1/users/by_email.json?email=${args.options.email}`
    }

    const requestOptions: any = {
      url: endPoint,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json;odata=nometadata'
      },
      json: true,
      body: {}
    };

    request
      .get(requestOptions)
      .then((res: any): void => {
        if (args.options.output === 'json') {
          cmd.log(res);
        }
        else {
          if (res instanceof Array)
            cmd.log((res as any[]).map(n => {
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
          else 
            cmd.log({ id: res.id, full_name: res.full_name, email: res.email, job_title: res.job_title, state: res.state, url: res.url });
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--id, --userId [number]',
        description: 'Retrieve a user by ID'
      },
      {
        option: '--email [string]',
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

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      ` Examples:
  
    Returns the current user
      ${this.name} user get

    Returns the user with the ID 1496550697
      ${this.name} user get --userId 1496550697

    Returns an array of users matching the e-mail john.smith@contoso.com
      ${this.name} user get --email john.smith@contoso.com

    Returns an array of users matching the e-mail john.smith@contoso.com in JSON. The JSON output returns a full user object.
      ${this.name} user get --email john.smith@contoso.com --output json
    `);
  }
}

module.exports = new YammerUserGetCommand();