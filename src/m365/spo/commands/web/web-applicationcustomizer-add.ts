import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { Options as spoCustomActionAddCommandOptions } from '../customaction/customaction-add';
import * as spoCustomActionAddCommand from '../customaction/customaction-add';
import Command from '../../../../Command';
import { Cli } from '../../../../cli/Cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  title: string;
  webUrl: string;
  clientSideComponentId: string;
  clientSideComponentProperties?: string;
}

class SpoWebApplicationCustomizerAddCommand extends SpoCommand {
  public get name(): string {
    return commands.WEB_APPLICATIONCUSTOMIZER_ADD;
  }

  public get description(): string {
    return 'Add an application customizer to a site.';
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
          return `${args.options.clientSideComponentId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        logger.logToStderr(`Adding application customizer with title '${args.options.title}' and clientSideComponentId '${args.options.clientSideComponentId}' to the site`);
      }

      const options: spoCustomActionAddCommandOptions = {
        webUrl: args.options.webUrl,
        name: args.options.title,
        title: args.options.title,
        clientSideComponentId: args.options.clientSideComponentId,
        clientSideComponentProperties: args.options.clientSideComponentProperties || '',
        location: 'ClientSideExtension.ApplicationCustomizer'
      };
      await Cli.executeCommand(spoCustomActionAddCommand as Command, { options: { ...options, _: [] } });
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoWebApplicationCustomizerAddCommand();