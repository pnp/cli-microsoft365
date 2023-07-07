import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { spo } from '../../../../utils/spo';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  confirm?: boolean;
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
        confirm: typeof args.options.confirm !== 'undefined'
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
        option: '--confirm'
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
    if (args.options.confirm) {
      await this.removeNode(logger, args.options);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the node ${args.options.id} from the navigation?`
      });

      if (result.continue) {
        await this.removeNode(logger, args.options);
      }
    }
  }

  private async removeNode(logger: Logger, options: Options): Promise<void> {
    try {
      const res = await spo.getRequestDigest(options.webUrl);

      if (this.verbose) {
        logger.logToStderr(`Removing navigation node...`);
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

module.exports = new SpoNavigationNodeRemoveCommand();