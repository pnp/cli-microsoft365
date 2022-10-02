import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { accessToken } from '../../../../utils/accessToken';
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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
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

    try {
      const res: any = await request.get(requestOptions);

      if (res.error) {
        throw res.error_description;
      }

      if (res.token_endpoint) {
        logger.log(res.token_endpoint.split('/')[3]);
      }
    } 
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new TenantIdGetCommand();