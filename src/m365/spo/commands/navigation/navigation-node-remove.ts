import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  force?: boolean;
  id: string;
  location: string;
  webUrl: string;
}

class SpoNavigationNodeRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.NAVIGATION_NODE_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified navigation node';
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
        location: args.options.location,
        force: typeof args.options.force !== 'undefined'
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
      },
      {
        option: '-i, --id <id>'
      },
      {
        option: '-f, --force'
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

        const id: number = parseInt(args.options.id);
        if (isNaN(id)) {
          return `${args.options.id} is not a number`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removeNode(logger, args.options);
    }
    else {
      const result = await Cli.promptForConfirmation(`Are you sure you want to remove the node ${args.options.id} from the navigation?`);

      if (result) {
        await this.removeNode(logger, args.options);
      }
    }
  }

  private async removeNode(logger: Logger, options: Options): Promise<void> {
    try {
      const res = await spo.getRequestDigest(options.webUrl);

      if (this.verbose) {
        await logger.logToStderr(`Removing navigation node...`);
      }

      const requestOptions: any = {
        url: `${options.webUrl}/_api/web/navigation/${options.location.toLowerCase()}/getbyid(${options.id})`,
        headers: {
          accept: 'application/json;odata=nometadata',
          'X-RequestDigest': res.FormDigestValue
        },
        responseType: 'json'
      };

      await request.delete(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoNavigationNodeRemoveCommand();