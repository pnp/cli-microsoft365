import { z } from 'zod';
import { zod } from '../../../../utils/zod.js';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

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
      .refine((options: Options) => !options.audienceIds || [options.audienceIds, options.audienceNames].filter(o => o !== undefined).length === 1, {
        message: 'Use one of the following options when specifying the audience name: audienceIds or audienceNames.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
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

      const isMultipleVivaConnectionsEnabled = await this.getIsMultipleVivaConnectionsFlightEnabled(spoAdminUrl, logger);
      const homeSiteCount = await this.getHomeSiteCount(spoAdminUrl, logger);

      const shouldUseSimpleSetSPHSite = homeSiteCount <= 1 &&
        this.vivaConnectionsOnlySpecified(args.options);

      if (shouldUseSimpleSetSPHSite) {
        requestOptions.url += '/SetSPHSite';
        requestOptions.data.sphSiteUrl = args.options.siteUrl;
        if (args.options.vivaConnectionsDefaultStart !== undefined) {
          requestOptions.data.vivaConnectionsDefaultStart = args.options.vivaConnectionsDefaultStart;
        }
      }
      else if (args.options.vivaConnectionsDefaultStart !== undefined ||
        args.options.draftMode !== undefined ||
        args.options.audienceIds !== undefined ||
        args.options.audienceNames !== undefined ||
        args.options.targetedLicenseType !== undefined) {
        if (isMultipleVivaConnectionsEnabled) {
          requestOptions.url += '/UpdateTargetedSite';
        }
        else {
          requestOptions.url += '/SetSPHSiteWithConfiguration';
        }
        requestOptions.data.siteUrl = args.options.siteUrl;
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
        requestOptions.data.configurationParam = configuration;
      }
      else {
        if (isMultipleVivaConnectionsEnabled) {
          requestOptions.url += '/UpdateTargetedSite';
          requestOptions.data.siteUrl = args.options.siteUrl;
        }
      }

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

  private async getHomeSiteCount(spoAdminUrl: string, logger: Logger): Promise<number> {
    try {
      if (this.verbose) {
        await logger.logToStderr('Retrieving current home site count...');
      }

      const requestOptions: CliRequestOptions = {
        url: `${spoAdminUrl}/_api/SPO.Tenant/GetTargetedSitesDetails`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const res = await request.get<{ value: any[] }>(requestOptions);
      const count = res.value ? res.value.length : 0;

      if (this.verbose) {
        await logger.logToStderr(`Current home site count: ${count}`);
      }

      return count;
    }
    catch (err: any) {
      if (this.verbose) {
        await logger.logToStderr(`Warning: Could not retrieve home site count. Defaulting to 0. Error: ${err.message}`);
      }
      return 0;
    }
  }

  private vivaConnectionsOnlySpecified(options: Options): boolean {
    const hasVivaConnections = options.vivaConnectionsDefaultStart !== undefined;

    // Check if only siteUrl or vivaConnectionsDefaultStart (or both) are specified, and no other options
    const otherOptions = [
      'draftMode',
      'audienceIds',
      'audienceNames',
      'targetedLicenseType',
      'order'
    ];

    const hasOtherOptions = otherOptions.some(opt => options[opt as keyof Options] !== undefined);

    return (hasVivaConnections) && !hasOtherOptions;
  }

  private async getIsMultipleVivaConnectionsFlightEnabled(spoAdminUrl: string, logger: Logger): Promise<boolean> {
    try {
      if (this.verbose) {
        await logger.logToStderr('Checking IsMultipleVivaConnectionsFlightEnabled tenant property...');
      }

      const requestOptions: CliRequestOptions = {
        url: `${spoAdminUrl}/_api/SPO.Tenant?$select=IsMultipleVivaConnectionsFlightEnabled`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const res = await request.get<{ IsMultipleVivaConnectionsFlightEnabled: boolean }>(requestOptions);
      return res.IsMultipleVivaConnectionsFlightEnabled;
    }
    catch (err: any) {
      if (this.verbose) {
        await logger.logToStderr(`Warning: Could not retrieve IsMultipleVivaConnectionsFlightEnabled property. Defaulting to false. Error: ${err.message}`);
      }
      return false;
    }
  }

  private async getGroupIdByName(groupName: string): Promise<string> {
    try {
      const requestOptions: CliRequestOptions = {
        url: `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${groupName.replace(/'/g, "''")}'&$select=id`,
        headers: {
          accept: 'application/json'
        },
        responseType: 'json'
      };

      const res = await request.get<{ value: { id: string }[] }>(requestOptions);

      if (res.value.length === 0) {
        throw new Error(`Group '${groupName}' not found`);
      }

      if (res.value.length > 1) {
        throw new Error(`Multiple groups found with name '${groupName}'. Please use group ID instead.`);
      }

      return res.value[0].id;
    }
    catch (err: any) {
      throw new Error(`Failed to get group ID for '${groupName}': ${err.message}`);
    }
  }

  private async transformAudienceNamesToIds(audienceNames: string): Promise<string[]> {
    const names = audienceNames.split(',').map(name => name.trim());
    const ids: string[] = [];

    for (const name of names) {
      const id = await this.getGroupIdByName(name);
      ids.push(id);
    }

    return ids;
  }
}
export default new SpoHomeSiteSetCommand();