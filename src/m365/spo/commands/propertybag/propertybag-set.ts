import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { IdentityResponse, spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import { SpoPropertyBagBaseCommand } from './propertybag-base.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  key: string;
  value: string;
  folder?: string;
}

class SpoPropertyBagSetCommand extends SpoPropertyBagBaseCommand {
  public get name(): string {
    return commands.PROPERTYBAG_SET;
  }

  public get description(): string {
    return 'Sets the value of the specified property in the property bag. Adds the property if it does not exist';
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
        option: '-v, --value <value>'
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

      let identityResp = await spo.getCurrentWebIdentity(args.options.webUrl, this.formDigestValue);
      const webIdentityResp = identityResp;

      // Check if web no script enabled or not
      // Cannot set property bag value if no script is enabled
      const isNoScriptSite = await this.isNoScriptSite(identityResp, args.options, logger);

      if (isNoScriptSite) {
        throw 'Site has NoScript enabled, and setting property bag values is not supported';
      }

      const opts: Options = args.options;
      if (opts.folder) {
        // get the folder guid instead of the web guid
        identityResp = await spo.getFolderIdentity(webIdentityResp.objectIdentity, opts.webUrl, opts.folder, this.formDigestValue);
      }

      await this.setProperty(identityResp, args.options, logger);
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }

  private setProperty(identityResp: IdentityResponse, options: Options, logger: Logger): Promise<any> {
    return SpoPropertyBagBaseCommand.setProperty(options.key, options.value, options.webUrl, this.formDigestValue, identityResp, logger, this.debug, options.folder);
  }

  private isNoScriptSite(webIdentityResp: IdentityResponse, options: Options, logger: Logger): Promise<boolean> {
    return SpoPropertyBagBaseCommand.isNoScriptSite(options.webUrl, this.formDigestValue, webIdentityResp, logger, this.debug);
  }
}

export default new SpoPropertyBagSetCommand();