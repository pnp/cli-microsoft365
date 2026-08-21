import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import config from '../../../../config.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { ClientSvcResponse, ClientSvcResponseContents, ContextInfo, spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  appCatalogUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .optional()
    .alias('u'),
  key: z.string().alias('k'),
  force: z.boolean().optional().alias('f')
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoStorageEntityRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.STORAGEENTITY_REMOVE;
  }

  public get description(): string {
    return 'Removes tenant property stored on the specified SharePoint Online app catalog';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removeTenantProperty(logger, args);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to delete the ${args.options.key} tenant property?` });

      if (result) {
        await this.removeTenantProperty(logger, args);
      }
    }
  }

  private async removeTenantProperty(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let appCatalogUrl = args.options.appCatalogUrl;

      if (!appCatalogUrl) {
        appCatalogUrl = await spo.getTenantAppCatalogUrl(logger, this.debug) as string;

        if (!appCatalogUrl) {
          throw 'Tenant app catalog URL not found. Specify the URL of the app catalog site using the appCatalogUrl option.';
        }
      }

      if (this.verbose) {
        await logger.logToStderr(`Removing tenant property ${args.options.key} from ${appCatalogUrl}...`);
      }

      const spoAdminUrl: string = await spo.getSpoAdminUrl(logger, this.debug);
      const digestInfo: ContextInfo = await spo.getRequestDigest(spoAdminUrl);

      const requestOptions: CliRequestOptions = {
        url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': digestInfo.FormDigestValue
        },
        data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="31" ObjectPathId="30" /><ObjectPath Id="33" ObjectPathId="32" /><ObjectPath Id="35" ObjectPathId="34" /><Method Name="RemoveStorageEntity" Id="36" ObjectPathId="34"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.key)}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="30" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="32" ParentId="30" Name="GetSiteByUrl"><Parameters><Parameter Type="String">${formatting.escapeXml(appCatalogUrl)}</Parameter></Parameters></Method><Property Id="34" ParentId="32" Name="RootWeb" /></ObjectPaths></Request>`
      };

      const processQuery: string = await request.post(requestOptions);
      const json: ClientSvcResponse = JSON.parse(processQuery);
      const response: ClientSvcResponseContents = json[0];

      if (response.ErrorInfo) {
        throw response.ErrorInfo.ErrorMessage;
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoStorageEntityRemoveCommand();