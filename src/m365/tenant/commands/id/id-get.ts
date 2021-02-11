import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, {
    CommandError, CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  domainName: string;
}

class TenantIdGetCommand extends Command {
  public get name(): string {
    return commands.TENANT_ID_GET;
  }

  public get description(): string {
    return 'Gets Microsoft 365 tenant ID for the specified domain';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    let domainName: string = args.options.domainName;
    if (!domainName) {
      const userName: string = Utils.getUserNameFromAccessToken(auth.service.accessTokens[auth.defaultResource].accessToken);
      domainName = userName.split('@')[1];
    }

    const requestOptions: any = {
      url: `https://login.windows.net/${domainName}/.well-known/openid-configuration`,
      headers: {
        'content-type': 'application/json',
        accept: 'application/json',
        'x-anonymous': true
      },
      responseType: 'json'
    };

    request
      .get(requestOptions)
      .then((res: any): void => {
        if (res.error) {
          cb(new CommandError(res.error_description));
          return;
        }

        if (res.token_endpoint) {
          logger.log(res.token_endpoint.split('/')[3]);
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-d, --domainName [domainName]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new TenantIdGetCommand();