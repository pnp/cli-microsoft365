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
  disableEnvironmentCreationByNonAdminusers?: boolean;
  disablePortalsCreationByNonAdminusers?: boolean;
  disableSurveyFeedback?: boolean;
  disableTrialEnvironmentCreationByNonAdminusers?: boolean;
  disableCapacityAllocationByEnvironmentAdmins?: boolean;
  disableSupportTicketsVisibleByAllUsers?: boolean;
  disableDocsSearch?: boolean;
  disableCommunitySearch?: boolean;
  disableBingVideoSearch?: boolean;
  shareWithColleaguesUserLimit?: string;
  disableShareWithEveryone?: boolean;
  enableGuestsToMake?: boolean;
  disableAdminDigest?: boolean;
  disableDeveloperEnvironmentCreationByNonAdminUsers?: boolean;
  disableBillingPolicyCreationByNonAdminUsers?: boolean;
  disableChampionsInvitationReachout?: boolean;
  disableSkillsMatchInvitationReachout?: boolean;
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
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        walkMeOptOut: !!args.options.walkMeOptOut,
        disableNPSCommentsReachout: !!args.options.disableNPSCommentsReachout,
        disableNewsletterSendout: !!args.options.disableNewsletterSendout,
        disableEnvironmentCreationByNonAdminusers: !!args.options.disableEnvironmentCreationByNonAdminusers,
        disablePortalsCreationByNonAdminusers: !!args.options.disablePortalsCreationByNonAdminusers,
        disableSurveyFeedback: !!args.options.disableSurveyFeedback,
        disableTrialEnvironmentCreationByNonAdminusers: !!args.options.disableTrialEnvironmentCreationByNonAdminusers,
        disableCapacityAllocationByEnvironmentAdmins: !!args.options.disableCapacityAllocationByEnvironmentAdmins,
        disableSupportTicketsVisibleByAllUsers: !!args.options.disableSupportTicketsVisibleByAllUsers,
        disableDocsSearch: !!args.options.disableDocsSearch,
        disableCommunitySearch: !!args.options.disableCommunitySearch,
        disableBingVideoSearch: !!args.options.disableBingVideoSearch,
        shareWithColleaguesUserLimit: typeof args.options.shareWithColleaguesUserLimit !== 'undefined',
        disableShareWithEveryone: !!args.options.disableShareWithEveryone,
        enableGuestsToMake: !!args.options.enableGuestsToMake,
        disableAdminDigest: !!args.options.disableAdminDigest,
        disableDeveloperEnvironmentCreationByNonAdminUsers: !!args.options.disableDeveloperEnvironmentCreationByNonAdminUsers,
        disableBillingPolicyCreationByNonAdminUsers: !!args.options.disableBillingPolicyCreationByNonAdminUsers,
        disableChampionsInvitationReachout: !!args.options.disableChampionsInvitationReachout,
        disableSkillsMatchInvitationReachout: !!args.options.disableSkillsMatchInvitationReachout
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
        option: '--disableEnvironmentCreationByNonAdminusers [disableEnvironmentCreationByNonAdminusers]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--disablePortalsCreationByNonAdminusers [disablePortalsCreationByNonAdminusers]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--disableSurveyFeedback [disableSurveyFeedback]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--disableTrialEnvironmentCreationByNonAdminusers [disableTrialEnvironmentCreationByNonAdminusers]',
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
        option: '--disableChampionsInvitationReachout [disableChampionsInvitationReachout]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--disableSkillsMatchInvitationReachout [disableSkillsMatchInvitationReachout]',
        autocomplete: ['true', 'false']
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {

        const regexNumber: RegExp = new RegExp(/\d/g);
        if (args.options.shareWithColleaguesUserLimit && !regexNumber.test(args.options.shareWithColleaguesUserLimit)) {
          return `'${args.options.shareWithColleaguesUserLimit}' is not a valid number`;
        }

        if (typeof args.options.walkMeOptOut === 'undefined' &&
          typeof args.options.disableNPSCommentsReachout === 'undefined' &&
          typeof args.options.disableNewsletterSendout === 'undefined' &&
          typeof args.options.disableEnvironmentCreationByNonAdminusers === 'undefined' &&
          typeof args.options.disablePortalsCreationByNonAdminusers === 'undefined' &&
          typeof args.options.disableSurveyFeedback === 'undefined' &&
          typeof args.options.disableTrialEnvironmentCreationByNonAdminusers === 'undefined' &&
          typeof args.options.disableCapacityAllocationByEnvironmentAdmins === 'undefined' &&
          typeof args.options.disableSupportTicketsVisibleByAllUsers === 'undefined' &&
          typeof args.options.disableDocsSearch === 'undefined' &&
          typeof args.options.disableCommunitySearch === 'undefined' &&
          typeof args.options.disableBingVideoSearch === 'undefined' &&
          !args.options.shareWithColleaguesUserLimit &&
          typeof args.options.disableShareWithEveryone === 'undefined' &&
          typeof args.options.enableGuestsToMake === 'undefined' &&
          typeof args.options.disableAdminDigest === 'undefined' &&
          typeof args.options.disableDeveloperEnvironmentCreationByNonAdminUsers === 'undefined' &&
          typeof args.options.disableBillingPolicyCreationByNonAdminUsers === 'undefined' &&
          typeof args.options.disableChampionsInvitationReachout === 'undefined' &&
          typeof args.options.disableSkillsMatchInvitationReachout === 'undefined') {
          return 'Specify at least one property to update';
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
          disableMembersIndicator: args.options.disableMembersIndicator
        },
        governance: {
          disableAdminDigest: args.options.disableAdminDigest,
          disableDeveloperEnvironmentCreationByNonAdminUsers: args.options.disableDeveloperEnvironmentCreationByNonAdminUsers
        },
        licensing: {
          disableBillingPolicyCreationByNonAdminUsers: args.options.disableBillingPolicyCreationByNonAdminUsers
        },
        champions: {
          disableChampionsInvitationReachout: args.options.disableChampionsInvitationReachout,
          disableSkillsMatchInvitationReachout: args.options.disableSkillsMatchInvitationReachout
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