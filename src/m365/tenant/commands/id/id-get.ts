import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Command, {
  CommandOption,
  CommandError
} from '../../../../Command';
import auth from '../../../../Auth';
import Utils from '../../../../Utils';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  domainName: string;
}

class TenantIdGetCommand extends Command {
  public get name(): string {
    return `${commands.TENANT_ID_GET}`;
  }

  public get description(): string {
    return 'Gets Microsoft 365 tenant ID for the specified domain';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    let domainName: string = args.options.domainName;
    if (!domainName) {
      const userName: string = Utils.getUserNameFromAccessToken(auth.service.accessTokens[auth.defaultResource].value);
      domainName = userName.split('@')[1];
    }

    const requestOptions: any = {
      url: `https://login.windows.net/${domainName}/.well-known/openid-configuration`,
      headers: {
        'content-type': 'application/json',
        accept: 'application/json',
        'x-anonymous': true
      },
      json: true
    };

    request
      .get(requestOptions)
      .then((res: any): void => {
        if (res.error) {
          cb(new CommandError(res.error_description));
          return;
        }

        if (res.token_endpoint) {
          cmd.log(res.token_endpoint.split('/')[3]);
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-d, --domainName [domainName]',
        description: 'The domain name for which to retrieve the Microsoft 365 tenant ID'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new TenantIdGetCommand();