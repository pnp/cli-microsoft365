import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import config from '../../../../config.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { ClientSvcResponse, ClientSvcResponseContents, spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  force?: boolean;
}

class SpoSiteUnarchiveCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_UNARCHIVE;
  }

  public get description(): string {
    return 'Unarchives a site collection';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --url <url>'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args) => validation.isValidSharePointUrl(args.options.url)
    );
  }

  #initTypes(): void {
    this.types.string.push('url');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {

    if (args.options.force) {
      await this.unarchiveSite(logger, args.options.url);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to unarchive the site ${args.options.url}?` });

      if (result) {
        await this.unarchiveSite(logger, args.options.url);
      }
    }
  }

  private async unarchiveSite(logger: Logger, url: string): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Unarchiving site ${url}...`);
    }

    try {
      const adminCenterUrl = await spo.getSpoAdminUrl(logger, this.debug);
      const requestDigest = await spo.getRequestDigest(adminCenterUrl);
      const requestData = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="UnarchiveSiteByUrl"><Parameters><Parameter Type="String">${url}</Parameter></Parameters></Method></ObjectPaths></Request>`;

      const requestOptions: CliRequestOptions = {
        url: `${adminCenterUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': requestDigest.FormDigestValue
        },
        data: requestData
      };

      const response: string = await request.post(requestOptions);
      const json: ClientSvcResponse = JSON.parse(response);
      const responseContent: ClientSvcResponseContents = json[0];

      if (responseContent.ErrorInfo) {
        throw responseContent.ErrorInfo.ErrorMessage;
      }
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }
}

export default new SpoSiteUnarchiveCommand();