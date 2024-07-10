import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import config from '../../../../config.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { ClientSvcResponse, ClientSvcResponseContents, ContextInfo, spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteUrl: string;
  force?: boolean;
}

class SpoSiteAppCatalogRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_APPCATALOG_REMOVE;
  }

  public get description(): string {
    return 'Removes site collection app catalog from the specified site';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --siteUrl <siteUrl>'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.siteUrl)
    );
  }

  #initTypes(): void {
    this.types.string.push('siteUrl');
  }


  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeSiteAppcatalog = async (): Promise<void> => {
      const url: string = args.options.siteUrl;

      if (this.verbose) {
        await logger.logToStderr(`Disabling site collection app catalog...`);
      }

      try {
        const spoAdminUrl: string = await spo.getSpoAdminUrl(logger, this.debug);
        const requestDigest: ContextInfo = await spo.getRequestDigest(spoAdminUrl);

        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': requestDigest.FormDigestValue
          },
          data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="50" ObjectPathId="49" /><ObjectPath Id="52" ObjectPathId="51" /><ObjectPath Id="54" ObjectPathId="53" /><ObjectPath Id="56" ObjectPathId="55" /><ObjectPath Id="58" ObjectPathId="57" /><Method Name="Remove" Id="59" ObjectPathId="57"><Parameters><Parameter Type="String">${formatting.escapeXml(url)}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="49" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="51" ParentId="49" Name="GetSiteByUrl"><Parameters><Parameter Type="String">${formatting.escapeXml(url)}</Parameter></Parameters></Method><Property Id="53" ParentId="51" Name="RootWeb" /><Property Id="55" ParentId="53" Name="TenantAppCatalog" /><Property Id="57" ParentId="55" Name="SiteCollectionAppCatalogsSites" /></ObjectPaths></Request>`
        };

        const res: string = await request.post(requestOptions);
        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          throw response.ErrorInfo.ErrorMessage;
        }
        else {
          if (this.verbose) {
            await logger.logToStderr('Site collection app catalog disabled');
          }
        }
      }
      catch (err: any) {
        this.handleRejectedPromise(err);
      }
    };

    if (args.options.force) {
      await removeSiteAppcatalog();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the app catalog from ${args.options.siteUrl}?` });

      if (result) {
        await removeSiteAppcatalog();
      }
    }
  }
}

export default new SpoSiteAppCatalogRemoveCommand();
