import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.folder = (!(!args.options.folder)).toString();
    return telemetryProps;
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

  public options(): CommandOption[] {
    const options: CommandOption[] = [
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
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (validation.isValidSharePointUrl(args.options.webUrl) !== true) {
      return 'Missing required option url';
    }

    if (!args.options.key) {
      return 'Missing required option key';
    }

    if (!args.options.value) {
      return 'Missing required option value';
    }

    return true;
  }
}

module.exports = new SpoPropertyBagSetCommand();