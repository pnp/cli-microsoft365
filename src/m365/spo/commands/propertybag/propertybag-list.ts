import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import { spo, validation } from '../../../../utils';
import commands from '../../commands';
import { Property, SpoPropertyBagBaseCommand } from './propertybag-base';

export interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  webUrl: string;
  folder?: string;
}

class SpoPropertyBagListCommand extends SpoPropertyBagBaseCommand {
  public get name(): string {
    return commands.PROPERTYBAG_LIST;
  }

  public get description(): string {
    return 'Gets property bag values';
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
        option: '-f, --folder [folder]'
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

      logger.log(this.formatOutput(propertyBagData));
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }

  /**
   * The property bag data returned from the client.svc/ProcessQuery response
   * has to be formatted before displayed since the key, value objects
   * carry extra information.
   * @param propertyBag client.svc property bag javascript object
   */
  private formatOutput(propertyBag: any): Property[] {
    const result: Property[] = [];
    const keys = Object.keys(propertyBag);

    for (let i = 0; i < keys.length; i++) {

      if (keys[i] === '_ObjectType_') {
        // this is system data, do not include it
        continue;
      }
      const formattedProp = this.formatProperty(keys[i], propertyBag[keys[i]]);
      result.push(formattedProp);
    }
    return result;
  }
}

module.exports = new SpoPropertyBagListCommand();