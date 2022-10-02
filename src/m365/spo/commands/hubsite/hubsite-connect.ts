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
  hubSiteId: string;
}

class SpoHubSiteConnectCommand extends SpoCommand {
  public get name(): string {
    return commands.HUBSITE_CONNECT;
  }

  public get description(): string {
    return 'Connects the specified site collection to the given hub site';
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
      },
      {
        option: '-i, --hubSiteId <hubSiteId>'
      }
    );
  }
  
  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.url);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }
    
        if (!validation.isValidGuid(args.options.hubSiteId)) {
          return `${args.options.hubSiteId} is not a valid GUID`;
        }
    
        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const res = await spo.getRequestDigest(args.options.url);
      const requestOptions: any = {
        url: `${args.options.url}/_api/site/JoinHubSite('${encodeURIComponent(args.options.hubSiteId)}')`,
        headers: {
          'X-RequestDigest': res.FormDigestValue,
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoHubSiteConnectCommand();