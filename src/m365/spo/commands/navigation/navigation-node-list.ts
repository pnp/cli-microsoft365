import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { NavigationNode } from './NavigationNode';

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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving navigation nodes...`);
    }

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web/navigation/${args.options.location.toLowerCase()}`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get<{ value: NavigationNode[] }>(requestOptions)
      .then((res: { value: NavigationNode[] }): void => {
        logger.log(res.value.map(n => {
          return {
            Id: n.Id,
            Title: n.Title,
            Url: n.Url
          };
        }));

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }
}

module.exports = new SpoNavigationNodeListCommand();