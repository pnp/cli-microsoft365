import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { TenantProperty } from './TenantProperty.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  key: z.string().alias('k'),
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

class SpoStorageEntityGetCommand extends SpoCommand {
  public get name(): string {
    return commands.STORAGEENTITY_GET;
  }

  public get description(): string {
    return 'Get details for the specified tenant property';
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

      const requestOptions: CliRequestOptions = {
        url: `${appCatalogUrl}/_api/web/GetStorageEntity('${formatting.encodeQueryParameter(args.options.key)}')`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const property: TenantProperty = await request.get(requestOptions);
      if (property["odata.null"] === true) {
        if (this.verbose) {
          await logger.logToStderr(`Property with key ${args.options.key} not found`);
        }
      }
      else {
        await logger.log({
          Key: args.options.key,
          Value: property.Value,
          Description: property.Description,
          Comment: property.Comment
        });
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoStorageEntityGetCommand();