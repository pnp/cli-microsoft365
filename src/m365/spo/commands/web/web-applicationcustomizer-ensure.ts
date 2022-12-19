import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { Cli } from '../../../../cli/Cli';
import Command from '../../../../Command';
import { Options as SpoCustomActionGetCommandOptions } from '../customaction/customaction-get';
import * as SpoCustomActionGetCommand from '../customaction/customaction-get';
import { Options as SpoCustomActionSetCommandOptions } from '../customaction/customaction-set';
import * as SpoCustomActionSetCommand from '../customaction/customaction-set';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  title: string;
  webUrl: string;
  clientSideComponentId: string;
  clientSideComponentProperties?: string;
}

class SpoWebApplicationCustomizerEnsureCommand extends SpoCommand {
  public get name(): string {
    return commands.WEB_APPLICATIONCUSTOMIZER_ENSURE;
  }

  public get description(): string {
    return 'Ensure an application customizer is added to a site.';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-t, --title <title>'
      },
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-i, --clientSideComponentId <clientSideComponentId>'
      },
      {
        option: '--clientSideComponentProperties [clientSideComponentProperties]'
      },
    );
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        clientSideComponentProperties: typeof args.options.clientSideComponentProperties !== 'undefined'
      });
    });
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.webUrl) {
          const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
          if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
          }
        }

        if (!validation.isValidGuid(args.options.clientSideComponentId)) {
          return `${args.options.clientSideComponentId} is not a valid clientSideComponentId`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        logger.logToStderr(`Ensuring if application customizer with id ${args.options.clientSideComponentId} exists in the site`);
      }

      const customActionId = await this.getCustomAction(args.options.webUrl, args.options.title, args.options.clientSideComponentId, logger);
      if (customActionId) {
        // Update
        logger.log('update item');
        await this.updateCustomAction(customActionId, args.options, logger);
      }
      else {
        // Add
        logger.log('add item');
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getCustomAction(webUrl: string, title: string, clientSideComponentId: string, logger: Logger): Promise<string> {
    if (this.verbose) {
      logger.logToStderr(`Checking if application customizer already exists on the site`);
    }
    const options: SpoCustomActionGetCommandOptions = {
      webUrl: webUrl,
      title: title,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    try {
      const output = await Cli.executeCommandWithOutput(SpoCustomActionGetCommand as Command, { options: { ...options, _: [] } });
      const getCustomActionOutput = JSON.parse(output.stdout);
      if (getCustomActionOutput.ClientSideComponentId !== clientSideComponentId) {
        throw `A custom component with the title ${title} was found, but the clientSideComponentId differs. Expected clientSideComponentId: ${getCustomActionOutput.ClientSideComponentId}`;
      }
      return getCustomActionOutput.Id;
    }
    catch (err: any) {
      if (err && err.error && err.error.message && err.error.message === `No user custom action with title '${title}' found`) {
        return '';
      }
      throw err;
    }
  }

  private async updateCustomAction(customActionId: string, options: Options, logger: Logger): Promise<void> {
    logger.log(customActionId);
    logger.log(options);
  }
}

module.exports = new SpoWebApplicationCustomizerEnsureCommand();