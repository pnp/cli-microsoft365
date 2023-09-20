import chalk from 'chalk';
import os from 'os';
import { Cli } from '../../cli/Cli.js';
import { Logger } from '../../cli/Logger.js';
import GlobalOptions from '../../GlobalOptions.js';
import { settingsNames } from '../../settingsNames.js';
import { CheckStatus, formatting } from '../../utils/formatting.js';
import { pid } from '../../utils/pid.js';
import AnonymousCommand from '../base/AnonymousCommand.js';
import commands from './commands.js';
import { interactivePreset, powerShellPreset, scriptingPreset } from './setupPresets.js';
import { prompt } from '../../utils/prompt.js';

interface Preferences {
  experience?: string;
  summary?: boolean;
  usageMode?: string;
  usedInPowerShell?: boolean;
}

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  interactive?: boolean;
  scripting?: boolean;
}

export type SettingNames = {
  [key in keyof typeof settingsNames]?: string | boolean;
};

class SetupCommand extends AnonymousCommand {
  // used for injecting answers from tests
  private answers: Preferences = {};

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
        scripting: args.options.scripting
      };

      Object.assign(this.telemetryProperties, properties);
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '--interactive' },
      { option: '--scripting' }
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

      await this.configureSettings(settings, true, logger);
      return;
    }

    await logger.logToStderr(`Welcome to the CLI for Microsoft 365 setup!`);
    await logger.logToStderr(`This command will guide you through the process of configuring the CLI for your needs.`);
    await logger.logToStderr(`Please, answer the following questions and we'll define a set of settings to best match how you intend to use the CLI.`);
    await logger.logToStderr('');

    const preferences: Preferences = await prompt.forInput([
      {
        type: 'list',
        name: 'usageMode',
        message: 'How do you plan to use the CLI?',
        choices: [
          'Interactively',
          'Scripting'
        ]
      },
      {
        type: 'confirm',
        name: 'usedInPowerShell',
        message: 'Are you going to use the CLI in PowerShell?',
        when: (answers: Preferences) => answers.usageMode === 'Scripting',
        default: pid.isPowerShell()
      },
      {
        type: 'list',
        name: 'experience',
        message: 'How experienced are you in using the CLI?',
        choices: [
          'Beginner',
          'Proficient'
        ]
      },
      {
        type: 'confirm',
        name: 'summary',
        // invoked by inquirer
        /* c8 ignore next 4 */
        message: (answers: Preferences) => {
          settings = this.getSettings(answers);
          return this.getSummaryMessage(settings);
        }
      }
    ], this.answers);

    if (preferences.summary) {
      // used only for testing. Normally, we'd get the settings from the answers
      /* c8 ignore next 3 */
      if (!settings) {
        settings = this.getSettings(preferences);
      }

      await logger.logToStderr('');
      await logger.logToStderr('Configuring settings...');
      await logger.logToStderr('');

      await this.configureSettings(settings, false, logger);

      if (!this.verbose) {
        await logger.logToStderr('');
        await logger.logToStderr(chalk.green('DONE'));
      }
    }
  }

  private getSummaryMessage(settings: SettingNames): string {
    const messageLines = [`Based on your preferences, we'll configure the following settings:`];
    for (const [key, value] of Object.entries(settings)) {
      messageLines.push(`- ${key}: ${value}`);
    }
    messageLines.push('', 'You can change any of these settings later using the `m365 cli config set` command or reset them to default using `m365 cli config reset`.', '', 'Do you want to apply these settings now?');

    return messageLines.join(os.EOL);
  }

  private getSettings(answers: Preferences): SettingNames {
    const settings: SettingNames = {};

    switch (answers.usageMode) {
      case 'Interactively':
        Object.assign(settings, interactivePreset);
        break;
      case 'Scripting':
        Object.assign(settings, scriptingPreset);
        break;
    }

    if (answers.usedInPowerShell === true) {
      Object.assign(settings, powerShellPreset);
    }

    switch (answers.experience) {
      case 'Beginner':
        settings.helpMode = 'full';
        break;
      case 'Proficient':
        settings.helpMode = 'options';
        break;
    }

    return settings;
  }

  private async configureSettings(settings: SettingNames, silent: boolean, logger: Logger): Promise<void> {
    if (this.debug) {
      await logger.logToStderr('Configuring settings...');
      await logger.logToStderr(JSON.stringify(settings, null, 2));
    }

    for (const [key, value] of Object.entries(settings)) {
      Cli.getInstance().config.set(key, value);

      if (!silent) {
        await logger.logToStderr(formatting.getStatus(CheckStatus.Success, `${key}: ${value}`));
      }
    }
  }
}

export default new SetupCommand();