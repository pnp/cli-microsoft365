import { z } from 'zod';
import { zod } from '../../../../utils/zod.js';
import { globalOptionsZod } from '../../../../Command.js';
import { validation } from '../../../../utils/validation.js';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { odata } from '../../../../utils/odata.js';
import config from '../../../../config.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { ClientSvcResponse, ClientSvcResponseContents, spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

const options = globalOptionsZod
  .extend({
    url: zod.alias('u', z.string()
      .refine((url: string) => validation.isValidSharePointUrl(url) === true, url => ({
        message: `'${url}' is not a valid SharePoint Online site URL.`
      }))
    ).optional(),
    force: zod.alias('f', z.boolean().optional())
  })
  .strict();

declare type Options = z.infer<typeof options>;
interface CommandArgs {
  options: Options;
}

class SpoHomeSiteRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.HOMESITE_REMOVE;
  }

  public get description(): string {
    return 'Removes a Home Site';
  }

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {

    const removeHomeSite: () => Promise<void> = async (): Promise<void> => {
      try {
        const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
        const reqDigest = await spo.getRequestDigest(spoAdminUrl);

        if (args.options.url) {
          await this.removeHomeSiteByUrl(args.options.url, spoAdminUrl, logger);
        }
        else {
          await this.showDeprecationWarning(
            logger,
            commands.HOMESITE_REMOVE,
            `${commands.HOMESITE_REMOVE} --url <url>`
          );

          const homeSites = await this.getHomeSites(spoAdminUrl);

          if (homeSites.length > 1) {
            await this.removeHomeSiteByUrl(homeSites[0].Url, spoAdminUrl, logger);
          }
          else {
            const requestOptions: CliRequestOptions = {
              url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
              headers: {
                'X-RequestDigest': reqDigest.FormDigestValue
              },
              data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="28" ObjectPathId="27" /><Method Name="RemoveSPHSite" Id="29" ObjectPathId="27" /></Actions><ObjectPaths><Constructor Id="27" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
            };

            const res = await request.post<string>(requestOptions);

            const json: ClientSvcResponse = JSON.parse(res);
            const response: ClientSvcResponseContents = json[0];
            if (response.ErrorInfo) {
              throw response.ErrorInfo.ErrorMessage;
            }
            else {
              await logger.log(json[json.length - 1]);
            }
          }
        }
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };


    if (args.options.force) {
      await removeHomeSite();
    }
    else {
      const result = await cli.promptForConfirmation({
        message: args.options.url
          ? `Are you sure you want to remove the Home Site at ${args.options.url}?`
          : `Are you sure you want to remove the primary Home Site?`
      });

      if (result) {
        await removeHomeSite();
      }
    }
  }

  private async removeHomeSiteByUrl(siteUrl: string, spoAdminUrl: string, logger: Logger): Promise<void> {
    const siteAdminProperties = await spo.getSiteAdminPropertiesByUrl(siteUrl, false, logger, this.verbose);

    const requestOptions: CliRequestOptions = {
      url: `${spoAdminUrl}/_api/SPO.Tenant/RemoveTargetedSite`,
      headers: {
        'Accept': 'application/json;odata=nometadata'
      },
      data: {
        siteId: siteAdminProperties.SiteId
      }
    };

    await request.post(requestOptions);

    await logger.log(`${siteUrl} has been removed as a Home Site. It may take some time for the change to apply. Check aka.ms/homesites for details.`);
  }

  private async getHomeSites(spoAdminUrl: string): Promise<any[]> {
    const requestOptions: CliRequestOptions = {
      url: `${spoAdminUrl}/_api/SPO.Tenant/GetTargetedSitesDetails`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return await odata.getAllItems(requestOptions);
  }
}

export default new SpoHomeSiteRemoveCommand();