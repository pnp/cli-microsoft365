import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import { Property, SpoPropertyBagBaseCommand } from './propertybag-base.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  key: string;
  folder?: string;
}

class SpoPropertyBagGetCommand extends SpoPropertyBagBaseCommand {
  public get name(): string {
    return commands.PROPERTYBAG_GET;
  }

  public get description(): string {
    return 'Gets the value of the specified property from the property bag';
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
        folder: typeof args.options.folder !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-k, --key <key>'
      },
      {
        option: '--folder [folder]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.webUrl)
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const contextResponse = await spo.getRequestDigest(args.options.webUrl);
      this.formDigestValue = contextResponse.FormDigestValue;

      const identityResp = await spo.getCurrentWebIdentity(args.options.webUrl, this.formDigestValue);

      let propertyBagData: any;
      const opts: Options = args.options;
      if (opts.folder) {
        propertyBagData = await this.getFolderPropertyBag(identityResp, opts.webUrl, opts.folder, logger);
      }
      else {
        propertyBagData = await this.getWebPropertyBag(identityResp, opts.webUrl, logger);
      }
      const property = this.filterByKey(propertyBagData, args.options.key);

      if (property) {
        await logger.log(property.value);
      }
      else if (this.verbose) {
        await logger.logToStderr('Property not found.');
      }
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }

  private filterByKey(propertyBag: any, key: string): Property | null {
    const keys = Object.keys(propertyBag);

    for (let i = 0; i < keys.length; i++) {
      // we have to normalize the keys and values before we can filter
      // since they carry extra information
      // ex. : 'vti_level$  Int32' instead of 'vti_level'
      const formattedProperty = this.formatProperty(keys[i], propertyBag[keys[i]]);
      if (formattedProperty.key === key) {
        return formattedProperty;
      }
    }

    return null;
  }
}

export default new SpoPropertyBagGetCommand();