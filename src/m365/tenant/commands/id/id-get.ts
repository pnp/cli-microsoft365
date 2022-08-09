import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, {
  CommandError
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { accessToken } from '../../../../utils';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  domainName?: string;
}

class TenantIdGetCommand extends Command {
  public get name(): string {
    return commands.ID_GET;
  }

  public get description(): string {
    return 'Gets Microsoft 365 tenant ID for the specified domain';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        domainName: typeof args.options.domainName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-d, --domainName [domainName]'
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    let domainName: string | undefined = args.options.domainName;
    if (!domainName) {
      const userName: string = accessToken.getUserNameFromAccessToken(auth.service.accessTokens[auth.defaultResource].accessToken);
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
}

module.exports = new TenantIdGetCommand();