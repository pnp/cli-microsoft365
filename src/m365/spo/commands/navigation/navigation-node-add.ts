import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  isExternal?: boolean;
  location?: string;
  parentNodeId?: number;
  title: string;
  url: string;
  webUrl: string;
}

class SpoNavigationNodeAddCommand extends SpoCommand {
  public get name(): string {
    return commands.NAVIGATION_NODE_ADD;
  }

  public get description(): string {
    return 'Adds a navigation node to the specified site navigation';
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
        isExternal: args.options.isExternal,
        parentNodeId: typeof args.options.parentNodeId !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-l, --location [location]',
        autocomplete: ['QuickLaunch', 'TopNavigationBar']
      },
      {
        option: '-t, --title <title>'
      },
      {
        option: '--url <url>'
      },
      {
        option: '--parentNodeId [parentNodeId]'
      },
      {
        option: '--isExternal'
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

        if (args.options.parentNodeId) {
          if (isNaN(args.options.parentNodeId)) {
            return `${args.options.parentNodeId} is not a number`;
          }
        }
        else {
          if (args.options.location !== 'QuickLaunch' &&
            args.options.location !== 'TopNavigationBar') {
            return `${args.options.location} is not a valid value for the location option. Allowed values are QuickLaunch|TopNavigationBar`;
          }
        }

        return true;
      }
    );
  }

  protected getExcludedOptionsWithUrls(): string[] | undefined {
    return ['url'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Adding navigation node...`);
    }

    const nodesCollection: string = args.options.parentNodeId ?
      `GetNodeById(${args.options.parentNodeId})/Children` :
      (args.options.location as string).toLowerCase();

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web/navigation/${nodesCollection}`,
      headers: {
        accept: 'application/json;odata=nometadata',
        'content-type': 'application/json;odata=nometadata'
      },
      data: {
        Title: args.options.title,
        Url: args.options.url,
        IsExternal: args.options.isExternal === true
      },
      responseType: 'json'
    };

    request
      .post(requestOptions)
      .then((res: any): void => {
        logger.log(res);
        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }
}

module.exports = new SpoNavigationNodeAddCommand();