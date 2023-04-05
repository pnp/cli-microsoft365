import GlobalOptions from '../../../../GlobalOptions';
import { Logger } from '../../../../cli/Logger';
import request, { CliRequestOptions } from '../../../../request';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  walkMeOptOut?: boolean;
  disableNPSCommentsReachout?: boolean;
  disableNewsletterSendout?: boolean;
  disableEnvironmentCreationByNonAdminUsers?: boolean;
  disablePortalsCreationByNonAdminUsers?: boolean;
  disableSurveyFeedback?: boolean;
  disableTrialEnvironmentCreationByNonAdminUsers?: boolean;
  disableCapacityAllocationByEnvironmentAdmins?: boolean;
  disableSupportTicketsVisibleByAllUsers?: boolean;
  disableDocsSearch?: boolean;
  disableCommunitySearch?: boolean;
  disableBingVideoSearch?: boolean;
  shareWithColleaguesUserLimit?: string;
  disableShareWithEveryone?: boolean;
  enableGuestsToMake?: boolean;
  disableMembersIndicator?: boolean;
  disableMakerMatch?: boolean;
  disablePreferredDataLocationForTeamsEnvironment?: boolean;
  disableAdminDigest?: boolean;
  disableDeveloperEnvironmentCreationByNonAdminUsers?: boolean;
  disableBillingPolicyCreationByNonAdminUsers?: boolean;
  storageCapacityConsumptionWarningThreshold?: string;
  disableChampionsInvitationReachout?: boolean;
  disableSkillsMatchInvitationReachout?: boolean;
  disableCopilot?: boolean;
  enableOpenAiBotPublishing?: boolean;
  enableModelDataSharing?: boolean;
}

class PpTenantSettingsSetCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.TENANT_SETTINGS_SET;
  }

  public get description(): string {
    return 'Sets the global Power Platform configuration of the tenant';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initTypes();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        walkMeOptOut: typeof args.options.walkMeOptOut !== 'undefined',
        disableNPSCommentsReachout: typeof args.options.disableNPSCommentsReachout !== 'undefined',
        disableNewsletterSendout: typeof args.options.disableNewsletterSendout !== 'undefined',
        disableEnvironmentCreationByNonAdminUsers: typeof args.options.disableEnvironmentCreationByNonAdminUsers !== 'undefined',
        disablePortalsCreationByNonAdminUsers: typeof args.options.disablePortalsCreationByNonAdminUsers !== 'undefined',
        disableSurveyFeedback: typeof args.options.disableSurveyFeedback !== 'undefined',
        disableTrialEnvironmentCreationByNonAdminUsers: typeof args.options.disableTrialEnvironmentCreationByNonAdminUsers !== 'undefined',
        disableCapacityAllocationByEnvironmentAdmins: typeof args.options.disableCapacityAllocationByEnvironmentAdmins !== 'undefined',
        disableSupportTicketsVisibleByAllUsers: typeof args.options.disableSupportTicketsVisibleByAllUsers !== 'undefined',
        disableDocsSearch: typeof args.options.disableDocsSearch !== 'undefined',
        disableCommunitySearch: typeof args.options.disableCommunitySearch !== 'undefined',
        disableBingVideoSearch: typeof args.options.disableBingVideoSearch !== 'undefined',
        shareWithColleaguesUserLimit: typeof args.options.shareWithColleaguesUserLimit !== 'undefined',
        disableShareWithEveryone: typeof args.options.disableShareWithEveryone !== 'undefined',
        enableGuestsToMake: typeof args.options.enableGuestsToMake !== 'undefined',
        disableMembersIndicator: typeof args.options.disableMembersIndicator !== 'undefined',
        disableMakerMatch: typeof args.options.disableMakerMatch !== 'undefined',
        disablePreferredDataLocationForTeamsEnvironment: typeof args.options.disablePreferredDataLocationForTeamsEnvironment !== 'undefined',
        disableAdminDigest: typeof args.options.disableAdminDigest !== 'undefined',
        disableDeveloperEnvironmentCreationByNonAdminUsers: typeof args.options.disableDeveloperEnvironmentCreationByNonAdminUsers !== 'undefined',
        disableBillingPolicyCreationByNonAdminUsers: typeof args.options.disableBillingPolicyCreationByNonAdminUsers !== 'undefined',
        storageCapacityConsumptionWarningThreshold: typeof args.options.storageCapacityConsumptionWarningThreshold !== 'undefined',
        disableChampionsInvitationReachout: typeof args.options.disableChampionsInvitationReachout !== 'undefined',
        disableSkillsMatchInvitationReachout: typeof args.options.disableSkillsMatchInvitationReachout !== 'undefined',
        disableCopilot: typeof args.options.disableCopilot !== 'undefined',
        enableOpenAiBotPublishing: typeof args.options.enableOpenAiBotPublishing !== 'undefined',
        enableModelDataSharing: typeof args.options.enableModelDataSharing !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--walkMeOptOut [walkMeOptOut]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--disableNPSCommentsReachout [disableNPSCommentsReachout]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--disableNewsletterSendout [disableNewsletterSendout]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--disableEnvironmentCreationByNonAdminUsers [disableEnvironmentCreationByNonAdminUsers]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--disablePortalsCreationByNonAdminUsers [disablePortalsCreationByNonAdminUsers]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--disableSurveyFeedback [disableSurveyFeedback]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--disableTrialEnvironmentCreationByNonAdminUsers [disableTrialEnvironmentCreationByNonAdminUsers]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--disableCapacityAllocationByEnvironmentAdmins [disableCapacityAllocationByEnvironmentAdmins]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--disableSupportTicketsVisibleByAllUsers [disableSupportTicketsVisibleByAllUsers]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--disableDocsSearch [disableDocsSearch]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--disableCommunitySearch [disableCommunitySearch]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--disableBingVideoSearch [disableBingVideoSearch]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--shareWithColleaguesUserLimit [shareWithColleaguesUserLimit]'
      },
      {
        option: '--disableShareWithEveryone [disableShareWithEveryone]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--enableGuestsToMake [enableGuestsToMake]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--disableMembersIndicator [disableMembersIndicator]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--disableMakerMatch [disableMakerMatch]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--disablePreferredDataLocationForTeamsEnvironment [disablePreferredDataLocationForTeamsEnvironment]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--disableAdminDigest [disableAdminDigest]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--disableDeveloperEnvironmentCreationByNonAdminUsers [disableDeveloperEnvironmentCreationByNonAdminUsers]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--disableBillingPolicyCreationByNonAdminUsers [disableBillingPolicyCreationByNonAdminUsers]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--storageCapacityConsumptionWarningThreshold [storageCapacityConsumptionWarningThreshold]'
      },
      {
        option: '--disableChampionsInvitationReachout [disableChampionsInvitationReachout]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--disableSkillsMatchInvitationReachout [disableSkillsMatchInvitationReachout]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--disableCopilot [disableCopilot]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--enableOpenAiBotPublishing [enableOpenAiBotPublishing]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--enableModelDataSharing [enableModelDataSharing]',
        autocomplete: ['true', 'false']
      }
    );
  }

  #initTypes(): void {
    this.types.boolean.push('walkMeOptOut', 'disableNPSCommentsReachout', 'disableNewsletterSendout', 'disableEnvironmentCreationByNonAdminUsers', 'disablePortalsCreationByNonAdminUsers', 'disableSurveyFeedback', 'disableTrialEnvironmentCreationByNonAdminUsers', 'disableCapacityAllocationByEnvironmentAdmins', 'disableSupportTicketsVisibleByAllUsers', 'disableDocsSearch', 'disableCommunitySearch', 'disableBingVideoSearch', 'disableShareWithEveryone', 'enableGuestsToMake', 'disableMembersIndicator', 'disableMakerMatch', 'disablePreferredDataLocationForTeamsEnvironment', 'disableAdminDigest', 'disableDeveloperEnvironmentCreationByNonAdminUsers', 'disableBillingPolicyCreationByNonAdminUsers', 'disableChampionsInvitationReachout', 'disableSkillsMatchInvitationReachout', 'disableCopilot', 'enableOpenAiBotPublishing', 'enableModelDataSharing');
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {

        if (args.options.shareWithColleaguesUserLimit) {
          const shareWithColleaguesUserLimit: number = parseInt(args.options.shareWithColleaguesUserLimit);
          if (isNaN(shareWithColleaguesUserLimit) || shareWithColleaguesUserLimit < 0) {
            return `'${args.options.shareWithColleaguesUserLimit}' is not a valid number`;
          }
        }

        if (args.options.storageCapacityConsumptionWarningThreshold) {
          const storageCapacityConsumptionWarningThreshold: number = parseInt(args.options.storageCapacityConsumptionWarningThreshold);
          if (isNaN(storageCapacityConsumptionWarningThreshold) || storageCapacityConsumptionWarningThreshold < 0) {
            return `'${args.options.storageCapacityConsumptionWarningThreshold}' is not a valid number`;
          }
        }

        if (typeof args.options.walkMeOptOut === 'undefined' &&
          typeof args.options.disableNPSCommentsReachout === 'undefined' &&
          typeof args.options.disableNewsletterSendout === 'undefined' &&
          typeof args.options.disableEnvironmentCreationByNonAdminUsers === 'undefined' &&
          typeof args.options.disablePortalsCreationByNonAdminUsers === 'undefined' &&
          typeof args.options.disableSurveyFeedback === 'undefined' &&
          typeof args.options.disableTrialEnvironmentCreationByNonAdminUsers === 'undefined' &&
          typeof args.options.disableCapacityAllocationByEnvironmentAdmins === 'undefined' &&
          typeof args.options.disableSupportTicketsVisibleByAllUsers === 'undefined' &&
          typeof args.options.disableDocsSearch === 'undefined' &&
          typeof args.options.disableCommunitySearch === 'undefined' &&
          typeof args.options.disableBingVideoSearch === 'undefined' &&
          !args.options.shareWithColleaguesUserLimit &&
          typeof args.options.disableShareWithEveryone === 'undefined' &&
          typeof args.options.enableGuestsToMake === 'undefined' &&
          typeof args.options.disableMembersIndicator === 'undefined' &&
          typeof args.options.disableMakerMatch === 'undefined' &&
          typeof args.options.disablePreferredDataLocationForTeamsEnvironment === 'undefined' &&
          typeof args.options.disableAdminDigest === 'undefined' &&
          typeof args.options.disableDeveloperEnvironmentCreationByNonAdminUsers === 'undefined' &&
          typeof args.options.disableBillingPolicyCreationByNonAdminUsers === 'undefined' &&
          !args.options.storageCapacityConsumptionWarningThreshold &&
          typeof args.options.disableChampionsInvitationReachout === 'undefined' &&
          typeof args.options.disableSkillsMatchInvitationReachout === 'undefined' &&
          typeof args.options.disableCopilot === 'undefined' &&
          typeof args.options.enableOpenAiBotPublishing === 'undefined' &&
          typeof args.options.enableModelDataSharing === 'undefined') {
          return 'Specify at least one option.';
        }

        return true;
      }
    );
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
          shareWithColleaguesUserLimit: args.options.shareWithColleaguesUserLimit
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
          storageCapacityConsumptionWarningThreshold: args.options.storageCapacityConsumptionWarningThreshold
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
      logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new PpTenantSettingsSetCommand();