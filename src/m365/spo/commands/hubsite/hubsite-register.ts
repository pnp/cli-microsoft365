import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { spo } from '../../../../utils/spo';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteUrl: string;
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
        option: '-u, --siteUrl <siteUrl>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.siteUrl)
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const reqDigest = await spo.getRequestDigest(args.options.siteUrl);

      const requestOptions: CliRequestOptions = {
        url: `${args.options.siteUrl}/_api/site/RegisterHubSite`,
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