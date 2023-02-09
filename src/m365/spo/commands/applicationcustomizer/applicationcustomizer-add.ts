import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import Command from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import * as spoCustomActionAddCommand from '../customaction/customaction-add';
import { Options as spoCustomActionAddCommandOptions } from '../customaction/customaction-add';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  title: string;
  webUrl: string;
  clientSideComponentId: string;
  clientSideComponentProperties?: string;
}

class SpoApplicationCustomizerAddCommand extends SpoCommand {
  public get name(): string {
    return commands.APPLICATIONCUSTOMIZER_ADD;
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

        if (args.options.clientSideComponentProperties) {
          try {
            JSON.parse(args.options.clientSideComponentProperties);
          }
          catch (e) {
            return `An error has occurred while parsing clientSideComponentProperties: ${e}`;
          }
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Adding application customizer with title '${args.options.title}' and clientSideComponentId '${args.options.clientSideComponentId}' to the site`);
    }

    const options: spoCustomActionAddCommandOptions = {
      webUrl: args.options.webUrl,
      name: args.options.title,
      title: args.options.title,
      clientSideComponentId: args.options.clientSideComponentId,
      clientSideComponentProperties: args.options.clientSideComponentProperties || '',
      location: 'ClientSideExtension.ApplicationCustomizer',
      debug: this.debug,
      verbose: this.verbose
    };
    await Cli.executeCommand(spoCustomActionAddCommand as Command, { options: { ...options, _: [] } });
  }
}

module.exports = new SpoApplicationCustomizerAddCommand();