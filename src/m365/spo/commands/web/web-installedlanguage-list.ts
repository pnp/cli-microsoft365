import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { WebInstalledLanguagePropertiesCollection } from './WebPropertiesCollection';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
}

class SpoWebInstalledLanguageListCommand extends SpoCommand {
  public get name(): string {
    return commands.WEB_INSTALLEDLANGUAGE_LIST;
  }

  public get description(): string {
    return 'Lists all installed languages on site';
  }

  public defaultProperties(): string[] | undefined {
    return ['DisplayName', 'LanguageTag', 'Lcid'];
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
      logger.logToStderr(`Retrieving all installed languages on site ${args.options.webUrl}...`);
    }

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web/RegionalSettings/InstalledLanguages`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get<WebInstalledLanguagePropertiesCollection>(requestOptions)
      .then((webInstalledLanguageProperties: WebInstalledLanguagePropertiesCollection): void => {
        logger.log(webInstalledLanguageProperties.Items);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new SpoWebInstalledLanguageListCommand();