import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils/validation';
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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
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

    try {
      const webInstalledLanguageProperties: WebInstalledLanguagePropertiesCollection = await request.get<WebInstalledLanguagePropertiesCollection>(requestOptions);
      logger.log(webInstalledLanguageProperties.Items);
    } 
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoWebInstalledLanguageListCommand();