import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { entraGroup } from '../../../../utils/entraGroup.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  url: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    }).alias('u'),
  vivaConnectionsDefaultStart: z.boolean().optional(),
  draftMode: z.boolean().optional(),
  audienceIds: z.string()
    .refine(audiences => validation.isValidGuidArray(audiences) === true, {
      error: e => `The following GUIDs are invalid: ${e.input}.`
    }).optional(),
  audienceNames: z.string().optional(),
  targetedLicenseType: z.enum(['everyone', 'frontLineWorkers', 'informationWorkers']).optional(),
  order: z.number()
    .refine(order => validation.isValidPositiveInteger(order) === true, {
      error: e => `'${e.input}' is not a positive integer.`
    }).optional()
});

type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoHomeSiteSetCommand extends SpoCommand {
  public get name(): string {
    return commands.HOMESITE_SET;
  }

  public get description(): string {
    return 'Updates an existing SharePoint home site.';
  }

  public get schema(): z.ZodType {
    return options;
  }


  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(
        (options: Options) => [options.audienceIds, options.audienceNames].filter(o => o !== undefined).length <= 1,
        {
          message: 'You must specify either audienceIds or audienceNames but not both.'
        }
      )
      .refine(
        (options: Options) =>
          options.vivaConnectionsDefaultStart !== undefined ||
          options.draftMode !== undefined ||
          options.audienceIds !== undefined ||
          options.audienceNames !== undefined ||
          options.targetedLicenseType !== undefined ||
          options.order !== undefined,
        {
          message: 'You must specify at least one option to configure.'
        }
      );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Configuring SharePoint home site: ${args.options.url}...`);
        await logger.logToStderr(`Attempting to retrieve the SharePoint admin URL.`);
      }

      const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);

      const configuration: any = {};
      if (args.options.vivaConnectionsDefaultStart !== undefined) {
        configuration.IsVivaConnectionsDefaultStartPresent = true;
        configuration.vivaConnectionsDefaultStart = args.options.vivaConnectionsDefaultStart;
      }
      if (args.options.draftMode !== undefined) {
        configuration.IsInDraftModePresent = true;
        configuration.isInDraftMode = args.options.draftMode;
      }
      if (args.options.audienceIds !== undefined) {
        configuration.IsAudiencesPresent = true;
        configuration.Audiences = args.options.audienceIds.split(',').map(id => id.trim());
      }
      if (args.options.audienceNames !== undefined) {
        configuration.IsAudiencesPresent = true;
        configuration.Audiences = args.options.audienceNames.trim() === '' ? [] : await this.transformAudienceNamesToIds(args.options.audienceNames);
      }
      if (args.options.targetedLicenseType !== undefined) {
        configuration.IsTargetedLicenseTypePresent = true;
        configuration.TargetedLicenseType = this.convertTargetedLicenseTypeToNumber(args.options.targetedLicenseType);
      }
      if (args.options.order !== undefined) {
        configuration.IsOrderPresent = true;
        configuration.Order = args.options.order;
      }

      const requestOptions: CliRequestOptions = {
        url: `${spoAdminUrl}/_api/SPO.Tenant/UpdateTargetedSite`,
        headers: {
          accept: 'application/json;odata=nometadata',
          'content-Type': 'application/json'
        },
        responseType: 'json',
        data: {
          siteUrl: args.options.url,
          configurationParam: configuration
        }
      };

      const res = await request.post(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private convertTargetedLicenseTypeToNumber(licenseType: string): number {
    const licenseTypeMap: Record<string, number> = {
      'everyone': 0,
      'frontLineWorkers': 1,
      'informationWorkers': 2
    };
    return licenseTypeMap[licenseType];
  }

  private async transformAudienceNamesToIds(audienceNames: string): Promise<string[]> {
    const names = audienceNames.split(',');
    const ids: string[] = [];

    for (const name of names) {
      const id = await entraGroup.getGroupIdByDisplayName(name.trim());
      ids.push(id);
    }

    return ids;
  }
}
export default new SpoHomeSiteSetCommand();