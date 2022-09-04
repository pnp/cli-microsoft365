import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { ContextInfo, spo, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  principals: string;
  rights: string;
}

class SpoSiteDesignRightsGrantCommand extends SpoCommand {
  public get name(): string {
    return commands.SITEDESIGN_RIGHTS_GRANT;
  }

  public get description(): string {
    return 'Grants access to a site design for one or more principals';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id <id>'
      },
      {
        option: '-p, --principals <principals>'
      },
      {
        option: '-r, --rights <rights>',
        autocomplete: ['View']
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        if (args.options.rights !== 'View') {
          return `${args.options.rights} is not a valid rights value. Allowed values View`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoUrl: string = await spo.getSpoUrl(logger, this.debug);
      const requestDigest: ContextInfo = await spo.getRequestDigest(spoUrl);
      const grantedRights: string = '1';
      const requestOptions: any = {
        url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GrantSiteDesignRights`,
        headers: {
          'X-RequestDigest': requestDigest.FormDigestValue,
          'content-type': 'application/json;charset=utf-8',
          accept: 'application/json;odata=nometadata'
        },
        data: {
          id: args.options.id,
          principalNames: args.options.principals.split(',').map(p => p.trim()),
          grantedRights: grantedRights
        },
        responseType: 'json'
      };

      return request.post(requestOptions);
    } 
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoSiteDesignRightsGrantCommand();
