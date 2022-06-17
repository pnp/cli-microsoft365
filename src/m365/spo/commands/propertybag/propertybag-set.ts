import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import { ContextInfo, IdentityResponse, spo, validation } from '../../../../utils';
import commands from '../../commands';
import { SpoPropertyBagBaseCommand } from './propertybag-base';

export interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
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
    let webIdentityResp: IdentityResponse;

    spo
      .getRequestDigest(args.options.webUrl)
      .then((contextResponse: ContextInfo): Promise<IdentityResponse> => {
        this.formDigestValue = contextResponse.FormDigestValue;

        return spo.getCurrentWebIdentity(args.options.webUrl, this.formDigestValue);
      })
      .then((identityResp: IdentityResponse): Promise<boolean> => {
        webIdentityResp = identityResp;

        // Check if web no script enabled or not
        // Cannot set property bag value if no script is enabled
        return this.isNoScriptSite(identityResp, args.options, logger);
      })
      .then((isNoScriptSite: boolean): Promise<IdentityResponse> => {
        if (isNoScriptSite) {
          return Promise.reject('Site has NoScript enabled, and setting property bag values is not supported');
        }

        const opts: Options = args.options;
        if (opts.folder) {
          // get the folder guid instead of the web guid
          return spo.getFolderIdentity(webIdentityResp.objectIdentity, opts.webUrl, opts.folder, this.formDigestValue);
        }

        return new Promise<IdentityResponse>(resolve => { return resolve(webIdentityResp); });
      })
      .then((identityResp: IdentityResponse): Promise<any> => {
        return this.setProperty(identityResp, args.options, logger);
      })
      .then(_ => cb(), (err: any): void => this.handleRejectedPromise(err, logger, cb));
  }

  private setProperty(identityResp: IdentityResponse, options: Options, logger: Logger): Promise<any> {
    return SpoPropertyBagBaseCommand.setProperty(options.key, options.value, options.webUrl, this.formDigestValue, identityResp, logger, this.debug, options.folder);
  }

  private isNoScriptSite(webIdentityResp: IdentityResponse, options: Options, logger: Logger): Promise<boolean> {
    return SpoPropertyBagBaseCommand.isNoScriptSite(options.webUrl, this.formDigestValue, webIdentityResp, logger, this.debug);
  }
}

module.exports = new SpoPropertyBagSetCommand();