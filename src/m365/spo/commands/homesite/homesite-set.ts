import { z } from 'zod';
import { zod } from '../../../../utils/zod.js';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { entraGroup } from '../../../../utils/entraGroup.js';

const optionsSchema = globalOptionsZod
  .extend({
    siteUrl: zod.alias('u', z.string()
      .refine((url: string) => validation.isValidSharePointUrl(url) === true, url => ({
        message: `'${url}' is not a valid SharePoint Online site URL.`
      }))
    ),
    vivaConnectionsDefaultStart: z.boolean().optional(),
    draftMode: z.boolean().optional(),
    audienceIds: z.string()
      .refine(audiences => validation.isValidGuidArray(audiences) === true, audiences => ({
        message: `The following GUIDs are invalid: ${validation.isValidGuidArray(audiences)}.`
      })).optional(),
    audienceNames: z.string().optional(),
    targetedLicenseType: z.enum(['everyone', 'frontLineWorkers', 'informationWorkers']).optional(),
    order: z.number()
      .refine(order => validation.isValidPositiveInteger(order) === true, order => ({
        message: `'${order}' is not a positive integer.`
      })).optional()
  });

type Options = z.infer<typeof optionsSchema>;

interface CommandArgs {
  options: Options;
}

class SpoHomeSiteSetCommand extends SpoCommand {
  public get name(): string {
    return commands.HOMESITE_SET;
  }

  public get description(): string {
    return 'Sets the specified site as the Home Site';
  }

  public get schema(): z.ZodTypeAny {
    return optionsSchema;
  }

  public getRefinedSchema(schema: z.ZodTypeAny): z.ZodEffects<any> | undefined {
    return schema
      .refine((options: Options) => [options.audienceIds, options.audienceNames].filter(o => o !== undefined).length <= 1, {
        message: 'Use one of the following options when specifying the audience name: audienceIds or audienceNames.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr('Will not set the home site anymore, only update existing ones...');
        await logger.logToStderr(`Setting the SharePoint home site to: ${args.options.siteUrl}...`);
        await logger.logToStderr('Attempting to retrieve the SharePoint admin URL.');
      }

      const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
      const requestOptions: CliRequestOptions = {
        url: `${spoAdminUrl}/_api/SPO.Tenant`,
        headers: {
          accept: 'application/json;odata=nometadata',
          'content-Type': 'application/json'
        },
        responseType: 'json',
        data: {}
      };

      await logger.logToStderr('DEPRECATION WARNING: Using \'spo homesite set\' to add new home sites is deprecated. Use \'spo homesite add\' instead to add a new home site.');

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
        configuration.Audiences = await this.transformAudienceNamesToIds(args.options.audienceNames);
      }
      if (args.options.targetedLicenseType !== undefined) {
        configuration.IsTargetedLicenseTypePresent = true;
        configuration.TargetedLicenseType = this.convertTargetedLicenseTypeToNumber(args.options.targetedLicenseType);
      }
      if (args.options.order !== undefined) {
        configuration.IsOrderPresent = true;
        configuration.Order = args.options.order;
      }
      requestOptions.url += '/UpdateTargetedSite';
      requestOptions.data.siteUrl = args.options.siteUrl;
      requestOptions.data.configurationParam = configuration;

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
    const result = licenseTypeMap[licenseType];
    return result;
  }

  private async transformAudienceNamesToIds(audienceNames: string): Promise<string[]> {
    const names = audienceNames.split(',').map(name => name.trim());
    const ids: string[] = [];

    for (const name of names) {
      const id = await entraGroup.getGroupIdByDisplayName(name);
      ids.push(id);
    }

    return ids;
  }
}
export default new SpoHomeSiteSetCommand();