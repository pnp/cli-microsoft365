import commands from '../../commands';
import SpoCommand from '../../../base/SpoCommand';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import { ContextInfo } from '../../spo';
import { SpoPropertyBagBaseCommand } from './propertybag-base';
import GlobalOptions from '../../../../GlobalOptions';
import { ClientSvc, IdentityResponse } from '../../ClientSvc';
import { CommandInstance } from '../../../../cli';

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
    return `${commands.PROPERTYBAG_SET}`;
  }

  public get description(): string {
    return 'Sets the value of the specified property in the property bag. Adds the property if it does not exist';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.folder = (!(!args.options.folder)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const clientSvcCommons: ClientSvc = new ClientSvc(cmd, this.debug);

    let webIdentityResp: IdentityResponse;

    this
      .getRequestDigest(args.options.webUrl)
      .then((contextResponse: ContextInfo): Promise<IdentityResponse> => {
        this.formDigestValue = contextResponse.FormDigestValue;

        return clientSvcCommons.getCurrentWebIdentity(args.options.webUrl, this.formDigestValue);
      })
      .then((identityResp: IdentityResponse): Promise<boolean> => {
        webIdentityResp = identityResp;

        // Check if web no script enabled or not
        // Cannot set property bag value if no script is enabled
        return this.isNoScriptSite(identityResp, args.options, clientSvcCommons);
      })
      .then((isNoScriptSite: boolean): Promise<IdentityResponse> => {
        if (isNoScriptSite) {
          return Promise.reject('Site has NoScript enabled, and setting property bag values is not supported');
        }

        const opts: Options = args.options;
        if (opts.folder) {
          // get the folder guid instead of the web guid
          return clientSvcCommons.getFolderIdentity(webIdentityResp.objectIdentity, opts.webUrl, opts.folder, this.formDigestValue);
        }

        return new Promise<IdentityResponse>(resolve => { return resolve(webIdentityResp); });
      })
      .then((identityResp: IdentityResponse): Promise<any> => {
        return this.setProperty(identityResp, args.options, cmd);
      })
      .then((res: any): void => {
        if (this.verbose) {
          cmd.log('DONE');
        }

        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  private setProperty(identityResp: IdentityResponse, options: Options, cmd: CommandInstance): Promise<any> {
    return SpoPropertyBagBaseCommand.setProperty(options.key, options.value, options.webUrl, this.formDigestValue, identityResp, cmd, this.debug, options.folder);
  }

  private isNoScriptSite(webIdentityResp: IdentityResponse, options: Options, clientSvcCommons: ClientSvc): Promise<boolean> {
    return SpoPropertyBagBaseCommand.isNoScriptSite(options.webUrl, this.formDigestValue, webIdentityResp, clientSvcCommons);
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'The URL of the site in which the property should be set'
      },
      {
        option: '-k, --key <key>',
        description: 'Key of the property to be set. Case-sensitive'
      },
      {
        option: '-v, --value <value>',
        description: 'Value of the property to be set'
      },
      {
        option: '-f, --folder [folder]',
        description: 'Site-relative URL of the folder on which the property should be set',
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (SpoCommand.isValidSharePointUrl(args.options.webUrl) !== true) {
        return 'Missing required option url';
      }

      if (!args.options.key) {
        return 'Missing required option key';
      }

      if (!args.options.value) {
        return 'Missing required option value';
      }

      return true;
    };
  }
}

module.exports = new SpoPropertyBagSetCommand();