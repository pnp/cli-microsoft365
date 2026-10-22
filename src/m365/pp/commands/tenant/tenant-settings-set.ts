import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  walkMeOptOut: z.boolean().optional(),
  disableNPSCommentsReachout: z.boolean().optional(),
  disableNewsletterSendout: z.boolean().optional(),
  disableEnvironmentCreationByNonAdminUsers: z.boolean().optional(),
  disablePortalsCreationByNonAdminUsers: z.boolean().optional(),
  disableSurveyFeedback: z.boolean().optional(),
  disableTrialEnvironmentCreationByNonAdminUsers: z.boolean().optional(),
  disableCapacityAllocationByEnvironmentAdmins: z.boolean().optional(),
  disableSupportTicketsVisibleByAllUsers: z.boolean().optional(),
  disableDocsSearch: z.boolean().optional(),
  disableCommunitySearch: z.boolean().optional(),
  disableBingVideoSearch: z.boolean().optional(),
  shareWithColleaguesUserLimit: z.string().refine(val => {
    const num = Number(val);
    return Number.isInteger(num) && num >= 0;
  }, {
    error: 'The value must be a non-negative integer.'
  }).optional(),
  disableShareWithEveryone: z.boolean().optional(),
  enableGuestsToMake: z.boolean().optional(),
  disableMembersIndicator: z.boolean().optional(),
  disableMakerMatch: z.boolean().optional(),
  disablePreferredDataLocationForTeamsEnvironment: z.boolean().optional(),
  disableAdminDigest: z.boolean().optional(),
  disableDeveloperEnvironmentCreationByNonAdminUsers: z.boolean().optional(),
  disableBillingPolicyCreationByNonAdminUsers: z.boolean().optional(),
  storageCapacityConsumptionWarningThreshold: z.string().refine(val => {
    const num = Number(val);
    return Number.isInteger(num) && num >= 0;
  }, {
    error: 'The value must be a non-negative integer.'
  }).optional(),
  disableChampionsInvitationReachout: z.boolean().optional(),
  disableSkillsMatchInvitationReachout: z.boolean().optional(),
  disableCopilot: z.boolean().optional(),
  enableOpenAiBotPublishing: z.boolean().optional(),
  enableModelDataSharing: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PpTenantSettingsSetCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.TENANT_SETTINGS_SET;
  }

  public get description(): string {
    return 'Sets the global Power Platform configuration of the tenant';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts =>
        opts.walkMeOptOut !== undefined ||
        opts.disableNPSCommentsReachout !== undefined ||
        opts.disableNewsletterSendout !== undefined ||
        opts.disableEnvironmentCreationByNonAdminUsers !== undefined ||
        opts.disablePortalsCreationByNonAdminUsers !== undefined ||
        opts.disableSurveyFeedback !== undefined ||
        opts.disableTrialEnvironmentCreationByNonAdminUsers !== undefined ||
        opts.disableCapacityAllocationByEnvironmentAdmins !== undefined ||
        opts.disableSupportTicketsVisibleByAllUsers !== undefined ||
        opts.disableDocsSearch !== undefined ||
        opts.disableCommunitySearch !== undefined ||
        opts.disableBingVideoSearch !== undefined ||
        opts.shareWithColleaguesUserLimit !== undefined ||
        opts.disableShareWithEveryone !== undefined ||
        opts.enableGuestsToMake !== undefined ||
        opts.disableMembersIndicator !== undefined ||
        opts.disableMakerMatch !== undefined ||
        opts.disablePreferredDataLocationForTeamsEnvironment !== undefined ||
        opts.disableAdminDigest !== undefined ||
        opts.disableDeveloperEnvironmentCreationByNonAdminUsers !== undefined ||
        opts.disableBillingPolicyCreationByNonAdminUsers !== undefined ||
        opts.storageCapacityConsumptionWarningThreshold !== undefined ||
        opts.disableChampionsInvitationReachout !== undefined ||
        opts.disableSkillsMatchInvitationReachout !== undefined ||
        opts.disableCopilot !== undefined ||
        opts.enableOpenAiBotPublishing !== undefined ||
        opts.enableModelDataSharing !== undefined, {
        error: 'Specify at least one option.',
        params: {
          customCode: 'required'
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const data = {
      walkMeOptOut: args.options.walkMeOptOut,
      disableNPSCommentsReachout: args.options.disableNPSCommentsReachout,
      disableNewsletterSendout: args.options.disableNewsletterSendout,
      disableEnvironmentCreationByNonAdminUsers: args.options.disableEnvironmentCreationByNonAdminUsers,
      disablePortalsCreationByNonAdminUsers: args.options.disablePortalsCreationByNonAdminUsers,
      disableSurveyFeedback: args.options.disableSurveyFeedback,
      disableTrialEnvironmentCreationByNonAdminUsers: args.options.disableTrialEnvironmentCreationByNonAdminUsers,
      disableCapacityAllocationByEnvironmentAdmins: args.options.disableCapacityAllocationByEnvironmentAdmins,
      disableSupportTicketsVisibleByAllUsers: args.options.disableSupportTicketsVisibleByAllUsers,
      powerPlatform: {
        search: {
          disableDocsSearch: args.options.disableDocsSearch,
          disableCommunitySearch: args.options.disableCommunitySearch,
          disableBingVideoSearch: args.options.disableBingVideoSearch
        },
        teamsIntegration: {
          shareWithColleaguesUserLimit: args.options.shareWithColleaguesUserLimit !== undefined ? Number(args.options.shareWithColleaguesUserLimit) : undefined
        },
        powerApps: {
          disableShareWithEveryone: args.options.disableShareWithEveryone,
          enableGuestsToMake: args.options.enableGuestsToMake,
          disableMembersIndicator: args.options.disableMembersIndicator,
          disableMakerMatch: args.options.disableMakerMatch
        },
        environments: {
          disablePreferredDataLocationForTeamsEnvironment: args.options.disablePreferredDataLocationForTeamsEnvironment
        },
        governance: {
          disableAdminDigest: args.options.disableAdminDigest,
          disableDeveloperEnvironmentCreationByNonAdminUsers: args.options.disableDeveloperEnvironmentCreationByNonAdminUsers
        },
        licensing: {
          disableBillingPolicyCreationByNonAdminUsers: args.options.disableBillingPolicyCreationByNonAdminUsers,
          storageCapacityConsumptionWarningThreshold: args.options.storageCapacityConsumptionWarningThreshold !== undefined ? Number(args.options.storageCapacityConsumptionWarningThreshold) : undefined
        },
        champions: {
          disableChampionsInvitationReachout: args.options.disableChampionsInvitationReachout,
          disableSkillsMatchInvitationReachout: args.options.disableSkillsMatchInvitationReachout
        },
        intelligence: {
          disableCopilot: args.options.disableCopilot,
          enableOpenAiBotPublishing: args.options.enableOpenAiBotPublishing
        },
        modelExperimentation: {
          enableModelDataSharing: args.options.enableModelDataSharing
        }
      }
    };

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/providers/Microsoft.BusinessAppPlatform/scopes/admin/updateTenantSettings?api-version=2020-10-01`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json',
      data: data
    };

    try {
      const res = await request.post<any>(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PpTenantSettingsSetCommand();