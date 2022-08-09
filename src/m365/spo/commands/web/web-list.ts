import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { WebPropertiesCollection } from "./WebPropertiesCollection";

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
}

class SpoWebListCommand extends SpoCommand {
  public get name(): string {
    return commands.WEB_LIST;
  }

  public get description(): string {
    return 'Lists subsites of the specified site';
  }

  public defaultProperties(): string[] | undefined {
    return ['Title', 'Url', 'Id'];
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.webUrl)
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving all webs in site at ${args.options.webUrl}...`);
    }

    let requestUrl: string = `${args.options.webUrl}/_api/web/webs`;

    if (args.options.output !== 'json') {
      requestUrl += '?$select=Title,Id,URL';
    }

    const requestOptions: any = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get<WebPropertiesCollection>(requestOptions)
      .then((webProperties: WebPropertiesCollection): void => {
        logger.log(webProperties.value);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new SpoWebListCommand();