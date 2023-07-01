import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import { ContextInfo, spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { SiteDesignPrincipal } from './SiteDesignPrincipal.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteDesignId: string;
}

class SpoSiteDesignRightsListCommand extends SpoCommand {
  public get name(): string {
    return commands.SITEDESIGN_RIGHTS_LIST;
  }

  public get description(): string {
    return 'Gets a list of principals that have access to a site design';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --siteDesignId <siteDesignId>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.siteDesignId)) {
          return `${args.options.siteDesignId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoUrl: string = await spo.getSpoUrl(logger, this.debug);
      const requestDigest: ContextInfo = await spo.getRequestDigest(spoUrl);
      const requestOptions: any = {
        url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRights`,
        headers: {
          'X-RequestDigest': requestDigest.FormDigestValue,
          'content-type': 'application/json;charset=utf-8',
          accept: 'application/json;odata=nometadata'
        },
        data: { id: args.options.siteDesignId },
        responseType: 'json'
      };

      const res: { value: SiteDesignPrincipal[] } = await request.post(requestOptions);
      await logger.log(res.value.map(p => {
        p.Rights = p.Rights === "1" ? "View" : p.Rights;
        return p;
      }));
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteDesignRightsListCommand();