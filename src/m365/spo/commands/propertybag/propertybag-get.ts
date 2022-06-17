import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import { ContextInfo, IdentityResponse, spo, validation } from '../../../../utils';
import commands from '../../commands';
import { Property, SpoPropertyBagBaseCommand } from './propertybag-base';

export interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
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
        option: '-f, --folder [folder]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.webUrl)
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    spo
      .getRequestDigest(args.options.webUrl)
      .then((contextResponse: ContextInfo): Promise<IdentityResponse> => {
        this.formDigestValue = contextResponse.FormDigestValue;

        return spo.getCurrentWebIdentity(args.options.webUrl, this.formDigestValue);
      })
      .then((identityResp: IdentityResponse): Promise<any> => {
        const opts: Options = args.options;
        if (opts.folder) {
          return this.getFolderPropertyBag(identityResp, opts.webUrl, opts.folder, logger);
        }

        return this.getWebPropertyBag(identityResp, opts.webUrl, logger);
      })
      .then((propertyBagData: any): void => {
        const property = this.filterByKey(propertyBagData, args.options.key);

        if (property) {
          logger.log(property.value);
        }
        else if (this.verbose) {
          logger.logToStderr('Property not found.');
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
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

module.exports = new SpoPropertyBagGetCommand();