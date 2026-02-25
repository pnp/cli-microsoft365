import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { TenantProperty } from './TenantProperty.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  appCatalogUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .optional()
    .alias('u')
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoStorageEntityListCommand extends SpoCommand {
  public get name(): string {
    return commands.STORAGEENTITY_LIST;
  }

  public get description(): string {
    return 'Lists tenant properties stored on the specified SharePoint Online app catalog';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let appCatalogUrl = args.options.appCatalogUrl;

      if (!appCatalogUrl) {
        appCatalogUrl = await spo.getTenantAppCatalogUrl(logger, this.debug) as string;

        if (!appCatalogUrl) {
          throw 'Tenant app catalog URL not found. Specify the URL of the app catalog site using the appCatalogUrl option.';
        }
      }

      if (this.verbose) {
        await logger.logToStderr(`Retrieving details for all tenant properties in ${appCatalogUrl}...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${appCatalogUrl}/_api/web/AllProperties?$select=storageentitiesindex`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const web: { storageentitiesindex?: string } = await request.get<{ storageentitiesindex?: string }>(requestOptions);
      if (!web.storageentitiesindex ||
        web.storageentitiesindex.trim().length === 0) {
        if (this.verbose) {
          await logger.logToStderr('No tenant properties found');
        }
      }
      else {
        const properties: { [key: string]: TenantProperty } = JSON.parse(web.storageentitiesindex);
        const keys: string[] = Object.keys(properties);
        if (keys.length === 0) {
          if (this.verbose) {
            await logger.logToStderr('No tenant properties found');
          }
        }
        else {
          await logger.log(keys.map((key: string): any => {
            const property: TenantProperty = properties[key];
            return {
              Key: key,
              Value: property.Value,
              Description: property.Description,
              Comment: property.Comment
            };
          }));
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoStorageEntityListCommand();