import chalk from 'chalk';
import os from 'os';
import auth, { AuthType } from '../../Auth.js';
import { cli } from '../../cli/cli.js';
import { Logger } from '../../cli/Logger.js';
import GlobalOptions from '../../GlobalOptions.js';
import { settingsNames } from '../../settingsNames.js';
import { accessToken } from '../../utils/accessToken.js';
import { AppCreationOptions, AppInfo, entraApp } from '../../utils/entraApp.js';
import { CheckStatus, formatting } from '../../utils/formatting.js';
import { pid } from '../../utils/pid.js';
import { ConfirmationConfig, SelectionConfig } from '../../utils/prompt.js';
import AnonymousCommand from '../base/AnonymousCommand.js';
import commands from './commands.js';
import { interactivePreset, powerShellPreset, scriptingPreset } from './setupPresets.js';

export interface Preferences {
  clientId?: string;
  tenantId?: string;
  clientSecret?: string;
  clientCertificateFile?: string;
  clientCertificateBase64Encoded?: string;
  clientCertificatePassword?: string;
  entraApp?: EntraAppConfig;
  experience?: CliExperience;
  newEntraAppScopes?: NewEntraAppScopes;
  summary?: boolean;
  usageMode?: CliUsageMode;
  usedInPowerShell?: boolean;
}

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  interactive?: boolean;
  scripting?: boolean;
  skipApp?: boolean;
}

export enum CliUsageMode {
  Interactively = 'interactively',
  Scripting = 'scripting'
}

export enum CliExperience {
  Beginner = 'beginner',
  Proficient = 'proficient'
}

export enum EntraAppConfig {
  Create = 'create',
  UseExisting = 'useExisting',
  Skip = 'skip'
}

export enum NewEntraAppScopes {
  Minimal = 'minimal',
  All = 'all'
}

export enum HelpMode {
  Full = 'full',
  Options = 'options'
}

const minimalScopes = [
  'https://graph.microsoft.com/User.Read'
];
const allScopes = [
  'https://graph.windows.net/Directory.AccessAsUser.All',
  'https://management.azure.com/user_impersonation',
  'https://admin.services.crm.dynamics.com/user_impersonation',
  'https://graph.microsoft.com/AppCatalog.ReadWrite.All',
  'https://graph.microsoft.com/AuditLog.Read.All',
  'https://graph.microsoft.com/Bookings.Read.All',
  'https://graph.microsoft.com/Calendars.Read',
  'https://graph.microsoft.com/ChannelMember.ReadWrite.All',
  'https://graph.microsoft.com/ChannelMessage.Read.All',
  'https://graph.microsoft.com/ChannelMessage.ReadWrite',
  'https://graph.microsoft.com/ChannelMessage.Send',
  'https://graph.microsoft.com/ChannelSettings.ReadWrite.All',
  'https://graph.microsoft.com/Chat.ReadWrite',
  'https://graph.microsoft.com/Directory.AccessAsUser.All',
  'https://graph.microsoft.com/Directory.ReadWrite.All',
  'https://graph.microsoft.com/ExternalConnection.ReadWrite.All',
  'https://graph.microsoft.com/ExternalItem.ReadWrite.All',
  'https://graph.microsoft.com/Group.ReadWrite.All',
  'https://graph.microsoft.com/IdentityProvider.ReadWrite.All',
  'https://graph.microsoft.com/InformationProtectionPolicy.Read',
  'https://graph.microsoft.com/Mail.Read.Shared',
  'https://graph.microsoft.com/Mail.ReadWrite',
  'https://graph.microsoft.com/Mail.Send',
  'https://graph.microsoft.com/Notes.ReadWrite.All',
  'https://graph.microsoft.com/OnlineMeetingArtifact.Read.All',
  'https://graph.microsoft.com/OnlineMeetings.ReadWrite',
  'https://graph.microsoft.com/OnlineMeetingTranscript.Read.All',
  'https://graph.microsoft.com/PeopleSettings.ReadWrite.All',
  'https://graph.microsoft.com/Place.Read.All',
  'https://graph.microsoft.com/Policy.Read.All',
  'https://graph.microsoft.com/RecordsManagement.ReadWrite.All',
  'https://graph.microsoft.com/Reports.Read.All',
  'https://graph.microsoft.com/RoleAssignmentSchedule.ReadWrite.Directory',
  'https://graph.microsoft.com/RoleEligibilitySchedule.Read.Directory',
  'https://graph.microsoft.com/SecurityEvents.Read.All',
  'https://graph.microsoft.com/ServiceHealth.Read.All',
  'https://graph.microsoft.com/ServiceMessage.Read.All',
  'https://graph.microsoft.com/ServiceMessageViewpoint.Write',
  'https://graph.microsoft.com/Sites.Read.All',
  'https://graph.microsoft.com/Tasks.ReadWrite',
  'https://graph.microsoft.com/Team.Create',
  'https://graph.microsoft.com/TeamMember.ReadWrite.All',
  'https://graph.microsoft.com/TeamsAppInstallation.ReadWriteForUser',
  'https://graph.microsoft.com/TeamSettings.ReadWrite.All',
  'https://graph.microsoft.com/TeamsTab.ReadWrite.All',
  'https://graph.microsoft.com/User.Invite.All',
  'https://manage.office.com/ActivityFeed.Read',
  'https://manage.office.com/ServiceHealth.Read',
  'https://analysis.windows.net/powerbi/api/Dataset.Read.All',
  'https://api.powerapps.com//User',
  'https://microsoft.sharepoint-df.com/AllSites.FullControl',
  'https://microsoft.sharepoint-df.com/TermStore.ReadWrite.All',
  'https://microsoft.sharepoint-df.com/User.ReadWrite.All'
];

export type SettingNames = {
  [key in keyof typeof settingsNames]?: string | boolean;
};

class SetupCommand extends AnonymousCommand {
  public get name(): string {
    return commands.SETUP;
  }

  public get description(): string {
    return 'Sets up CLI for Microsoft 365 based on your preferences';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      const properties: any = {
        interactive: args.options.interactive,
        scripting: args.options.scripting,
        skipApp: args.options.skipApp
      };

      Object.assign(this.telemetryProperties, properties);
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '--interactive' },
      { option: '--scripting' },
      { option: '--skipApp' }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.interactive && args.options.scripting) {
          return 'Specify either interactive or scripting but not both';
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let settings: SettingNames | undefined;

    if (args.options.interactive || args.options.scripting) {
      settings = {};

      if (args.options.interactive) {
        Object.assign(settings, interactivePreset);
      }
      else if (args.options.scripting) {
        Object.assign(settings, scriptingPreset);
      }

      if (pid.isPowerShell()) {
        Object.assign(settings, powerShellPreset);
      }

      await this.configureSettings({ preferences: {}, settings, silent: true, logger });
      return;
    }

    await logger.logToStderr(`Welcome to the CLI for Microsoft 365 setup!`);
    await logger.logToStderr(`This command will guide you through the process of configuring the CLI for your needs.`);
    await logger.logToStderr(`Please, answer the following questions and we'll define a set of settings to best match how you intend to use the CLI.`);
    await logger.logToStderr('');

    const preferences: Preferences = {};

    if (!args.options.skipApp) {
      const entraAppConfig: SelectionConfig<EntraAppConfig> = {
        message: 'CLI for Microsoft 365 requires a Microsoft Entra app. Do you want to create a new app registration or use an existing one?',
        choices: [
          { name: 'Create a new app registration', value: EntraAppConfig.Create },
          { name: 'Use an existing app registration', value: EntraAppConfig.UseExisting },
          { name: 'Skip configuring app registration', value: EntraAppConfig.Skip }
        ]
      };
      preferences.entraApp = await cli.promptForSelection(entraAppConfig);
      switch (preferences.entraApp) {
        case EntraAppConfig.Create:
          const newEntraAppScopesConfig: SelectionConfig<NewEntraAppScopes> = {
            message: 'What scopes should the new app registration have?',
            choices: [
              { name: 'User.Read (you will need to add the necessary permissions yourself)', value: NewEntraAppScopes.Minimal },
              { name: 'All (easy way to use all CLI commands)', value: NewEntraAppScopes.All }
            ]
          };
          preferences.newEntraAppScopes = await cli.promptForSelection(newEntraAppScopesConfig);
          break;
        case EntraAppConfig.UseExisting:
          const existingApp = await this.configureExistingEntraApp(logger);
          Object.assign(preferences, existingApp);
          break;
      }
    }
    else {
      preferences.entraApp = EntraAppConfig.Skip;
    }

    const usageModeConfig: SelectionConfig<CliUsageMode> = {
      message: 'How do you plan to use the CLI?',
      choices: [
        { name: 'Interactively', value: CliUsageMode.Interactively },
        { name: 'Scripting', value: CliUsageMode.Scripting }
      ]
    };
    preferences.usageMode = await cli.promptForSelection(usageModeConfig);

    if (preferences.usageMode === CliUsageMode.Scripting) {
      const usedInPowerShellConfig: ConfirmationConfig = {
        message: 'Are you going to use the CLI in PowerShell?',
        default: pid.isPowerShell()
      };
      preferences.usedInPowerShell = await cli.promptForConfirmation(usedInPowerShellConfig);
    }

    const experienceConfig: SelectionConfig<CliExperience> = {
      message: 'How experienced are you in using the CLI?',
      choices: [
        { name: 'Beginner', value: CliExperience.Beginner },
        { name: 'Proficient', value: CliExperience.Proficient }
      ]
    };
    preferences.experience = await cli.promptForSelection(experienceConfig);

    const summaryConfig: ConfirmationConfig = {
      message: this.getSummaryMessage(preferences)
    };
    preferences.summary = await cli.promptForConfirmation(summaryConfig);

    if (!preferences.summary) {
      return;
    }

    // used only for testing. Normally, we'd get the settings from the answers
    /* c8 ignore next 3 */
    if (!settings) {
      settings = this.getSettings(preferences);
    }

    await logger.logToStderr('');
    await logger.logToStderr('Configuring settings...');
    await logger.logToStderr('');

    await this.configureSettings({ preferences, settings, silent: false, logger });

    if (!this.verbose) {
      await logger.logToStderr('');
      await logger.logToStderr(chalk.green('DONE'));
    }
  }

  private async configureExistingEntraApp(logger: Logger): Promise<Preferences> {
    await logger.logToStderr('Please provide the details of the existing app registration.');
    let clientCertificateFile: string | undefined;
    let clientCertificateBase64Encoded: string | undefined;
    let clientCertificatePassword: string | undefined;

    const clientId = await cli.promptForInput({
      message: 'Client ID:',
      /* c8 ignore next */
      validate: value => value.length > 0
    });
    const tenantId = await cli.promptForInput({
      message: 'Tenant ID (leave common if the app is multitenant):',
      default: 'common'
    });
    const clientSecret = await cli.promptForInput({
      message: 'Client secret (leave empty if you use a certificate or a public client):'
    });
    if (!clientSecret) {
      clientCertificateFile = await cli.promptForInput({
        message: `Path to the client certificate file (leave empty if you want to specify a base64-encoded certificate string):`
      });
      if (!clientCertificateFile) {
        clientCertificateBase64Encoded = await cli.promptForInput({
          message: 'Base64-encoded certificate string:'
        });
      }
      if (clientCertificateFile || clientCertificateBase64Encoded) {
        clientCertificatePassword = await cli.promptForInput({
          message: 'Password for the client certificate (leave empty if the certificate is not password-protected):'
        });
      }
    }

    return {
      clientId,
      tenantId,
      clientSecret,
      clientCertificateFile,
      clientCertificateBase64Encoded,
      clientCertificatePassword
    };
  }

  private async createNewEntraApp(preferences: Preferences, logger: Logger): Promise<AppInfo> {
    if (!await cli.promptForConfirmation({
      message: 'CLI for Microsoft 365 will now sign in to your Microsoft 365 tenant as Microsoft Azure CLI to create a new app registration. Continue?',
      default: false
    })) {
      throw new Error();
    }

    // logout, just in case
    await auth.clearConnectionInfo();
    // setup auth
    auth.connection.authType = AuthType.Browser;
    // Microsoft Azure CLI app ID
    auth.connection.appId = '04b07795-8ddb-461a-bbee-02f9e1bf7b46';
    auth.connection.tenant = 'common';
    await auth.ensureAccessToken(auth.defaultResource, logger, this.debug);
    auth.connection.active = true;

    const options: AppCreationOptions = {
      allowPublicClientFlows: true,
      apisDelegated: (preferences.newEntraAppScopes === NewEntraAppScopes.All ? allScopes : minimalScopes).join(','),
      implicitFlow: false,
      multitenant: false,
      name: 'CLI for Microsoft 365',
      platform: 'publicClient',
      redirectUris: 'http://localhost,https://localhost,https://login.microsoftonline.com/common/oauth2/nativeclient'
    };
    const apis = await entraApp.resolveApis({
      options,
      logger,
      verbose: this.verbose,
      debug: this.debug
    });
    const appInfo: AppInfo = await entraApp.createAppRegistration({
      options,
      apis,
      logger,
      verbose: this.verbose,
      debug: this.debug
    });
    appInfo.tenantId = accessToken.getTenantIdFromAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken);
    await entraApp.grantAdminConsent({
      appInfo,
      appPermissions: entraApp.appPermissions,
      adminConsent: true,
      logger,
      debug: this.debug
    });

    // logout
    await auth.clearConnectionInfo();

    return appInfo;
  }

  private getSummaryMessage(preferences: Preferences): string {
    const messageLines = [`Based on your preferences, we'll configure the following settings:`];
    switch (preferences.entraApp) {
      case EntraAppConfig.Create:
        messageLines.push(`- Entra app: Create a new app registration with ${preferences.newEntraAppScopes} scopes`);
        break;
      case EntraAppConfig.UseExisting:
        messageLines.push(`- Entra app: use existing`);
        messageLines.push(`  - Client ID: ${preferences.clientId}`);
        messageLines.push(`  - Tenant ID: ${preferences.tenantId}`);
        if (preferences.clientSecret) {
          messageLines.push(`  - Client secret: ${preferences.clientSecret}`);
        }
        if (preferences.clientCertificateFile) {
          messageLines.push(`  - Client certificate file: ${preferences.clientCertificateFile}`);
        }
        if (preferences.clientCertificateBase64Encoded) {
          messageLines.push(`  - Client certificate base64-encoded: ${preferences.clientCertificateBase64Encoded}`);
        }
        if (preferences.clientCertificatePassword) {
          messageLines.push(`  - Client certificate password: ${preferences.clientCertificatePassword}`);
        }
        break;
      case EntraAppConfig.Skip:
        messageLines.push(`- Entra app: skip`);
        break;
    }

    const settings: SettingNames = this.getSettings(preferences);
    for (const [key, value] of Object.entries(settings)) {
      messageLines.push(`- ${key}: ${value}`);
    }
    messageLines.push('', 'You can change any of these settings later using the `m365 cli config set` command or reset them to default using `m365 cli config reset`.', '', 'Do you want to apply these settings now?');

    return messageLines.join(os.EOL);
  }

  private getSettings(answers: Preferences): SettingNames {
    const settings: SettingNames = {};

    switch (answers.usageMode) {
      case CliUsageMode.Interactively:
        Object.assign(settings, interactivePreset);
        break;
      case CliUsageMode.Scripting:
        Object.assign(settings, scriptingPreset);
        break;
    }

    if (answers.usedInPowerShell === true) {
      Object.assign(settings, powerShellPreset);
    }

    switch (answers.experience) {
      case CliExperience.Beginner:
        settings.helpMode = HelpMode.Full;
        break;
      case CliExperience.Proficient:
        settings.helpMode = HelpMode.Options;
        break;
    }

    return settings;
  }

  private async configureSettings({ preferences, settings, silent, logger }: {
    preferences: Preferences,
    settings: SettingNames,
    silent: boolean,
    logger: Logger
  }): Promise<void> {
    switch (preferences.entraApp) {
      case EntraAppConfig.Create:
        if (this.verbose) {
          await logger.logToStderr('Creating a new Entra app...');
        }
        const appSettings = await this.createNewEntraApp(preferences, logger);
        Object.assign(settings, {
          clientId: appSettings.appId,
          tenantId: appSettings.tenantId
        });
        cli.getConfig().delete(settingsNames.clientSecret);
        cli.getConfig().delete(settingsNames.clientCertificateFile);
        cli.getConfig().delete(settingsNames.clientCertificateBase64Encoded);
        cli.getConfig().delete(settingsNames.clientCertificatePassword);
        break;
      case EntraAppConfig.UseExisting:
        Object.assign(settings, {
          clientId: preferences.clientId,
          tenantId: preferences.tenantId,
          clientSecret: preferences.clientSecret,
          clientCertificateFile: preferences.clientCertificateFile,
          clientCertificateBase64Encoded: preferences.clientCertificateBase64Encoded,
          clientCertificatePassword: preferences.clientCertificatePassword
        });
        break;
      case EntraAppConfig.Skip:
        break;
    }

    if (this.debug) {
      await logger.logToStderr('Configuring settings...');
      await logger.logToStderr(JSON.stringify(settings, null, 2));
    }

    for (const [key, value] of Object.entries(settings)) {
      cli.getConfig().set(key, value);

      if (!silent) {
        await logger.logToStderr(formatting.getStatus(CheckStatus.Success, `${key}: ${value}`));
      }
    }
  }
}

export default new SetupCommand();