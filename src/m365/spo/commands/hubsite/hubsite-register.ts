import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { spo } from '../../../../utils/spo';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
}

class SpoHubSiteRegisterCommand extends SpoCommand {
  public get name(): string {
    return commands.HUBSITE_REGISTER;
  }

  public get description(): string {
    return 'Registers the specified site collection as a hub site';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --url <url>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.url)
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const reqDigest = await spo.getRequestDigest(args.options.url);

      const requestOptions: any = {
        url: `${args.options.url}/_api/site/RegisterHubSite`,
        headers: {
          'X-RequestDigest': reqDigest.FormDigestValue,
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const res = await request.post(requestOptions);
      logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoHubSiteRegisterCommand();