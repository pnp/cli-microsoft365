import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { NavigationNode } from './NavigationNode.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  location: string;
  webUrl: string;
}

class SpoNavigationNodeListCommand extends SpoCommand {
  public get name(): string {
    return commands.NAVIGATION_NODE_LIST;
  }

  public get description(): string {
    return 'Lists nodes from the specified site navigation';
  }

  public defaultProperties(): string[] | undefined {
    return ['Id', 'Title', 'Url'];
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
        location: args.options.location
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-l, --location <location>',
        autocomplete: ['QuickLaunch', 'TopNavigationBar']
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.location !== 'QuickLaunch' &&
          args.options.location !== 'TopNavigationBar') {
          return `${args.options.location} is not a valid value for the location option. Allowed values are QuickLaunch|TopNavigationBar`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving navigation nodes...`);
    }

    try {
      const res = await odata.getAllItems<NavigationNode>(`${args.options.webUrl}/_api/web/navigation/${args.options.location.toLowerCase()}`);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoNavigationNodeListCommand();