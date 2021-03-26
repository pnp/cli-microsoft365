import { Cli, Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import SpoCommand from '../../../base/SpoCommand';
import { ClientSvc, IdentityResponse } from '../../ClientSvc';
import commands from '../../commands';
import { ClientSvcResponse, ClientSvcResponseContents, ContextInfo } from '../../spo';
import { SpoPropertyBagBaseCommand } from './propertybag-base';

export interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  webUrl: string;
  key: string;
  folder?: string;
  confirm?: boolean;
}

class SpoPropertyBagRemoveCommand extends SpoPropertyBagBaseCommand {
  public get name(): string {
    return commands.PROPERTYBAG_REMOVE;
  }

  public get description(): string {
    return 'Removes specified property from the property bag';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.folder = (!(!args.options.folder)).toString();
    telemetryProps.confirm = args.options.confirm === true;
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const removeProperty = (): void => {
      const clientSvcCommons: ClientSvc = new ClientSvc(logger, this.debug);

      this
        .getRequestDigest(args.options.webUrl)
        .then((contextResponse: ContextInfo): Promise<IdentityResponse> => {
          this.formDigestValue = contextResponse.FormDigestValue;

          return clientSvcCommons.getCurrentWebIdentity(args.options.webUrl, this.formDigestValue);
        })
        .then((identityResp: IdentityResponse): Promise<IdentityResponse> => {
          const opts: Options = args.options;
          if (opts.folder) {
            // get the folder guid instead of the web guid
            return clientSvcCommons.getFolderIdentity(identityResp.objectIdentity, opts.webUrl, opts.folder, this.formDigestValue);
          }
          return new Promise<IdentityResponse>(resolve => { return resolve(identityResp); });
        })
        .then((identityResp: IdentityResponse): Promise<any> => {
          return this.removeProperty(identityResp, args.options);
        })
        .then(_ => cb(), (err: any): void => this.handleRejectedPromise(err, logger, cb));
    };

    if (args.options.confirm) {
      removeProperty();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the ${args.options.key} property?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeProperty();
        }
      });
    }
  }

  private removeProperty(identityResp: IdentityResponse, options: Options): Promise<any> {
    let objectType: string = 'AllProperties';
    if (options.folder) {
      objectType = 'Properties';
    }

    const requestOptions: any = {
      url: `${options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': this.formDigestValue
      },
      data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetFieldValue" Id="206" ObjectPathId="205"><Parameters><Parameter Type="String">${Utils.escapeXml(options.key)}</Parameter><Parameter Type="Null" /></Parameters></Method><Method Name="Update" Id="207" ObjectPathId="198" /></Actions><ObjectPaths><Property Id="205" ParentId="198" Name="${objectType}" /><Identity Id="198" Name="${identityResp.objectIdentity}" /></ObjectPaths></Request>`
    };

    return new Promise<any>((resolve: any, reject: any): void => {
      request.post(requestOptions).then((res: any): void => {
        const json: ClientSvcResponse = JSON.parse(res);
        const contents: ClientSvcResponseContents = json.find(x => { return x['ErrorInfo']; });
        if (contents && contents.ErrorInfo) {
          reject(contents.ErrorInfo.ErrorMessage || 'ClientSvc unknown error');
        }
        else {
          resolve(res);
        }
      }, (err: any): void => { reject(err); });
    });
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
        option: '-f, --folder [folder]'
      },
      {
        option: '--confirm'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (SpoCommand.isValidSharePointUrl(args.options.webUrl) !== true) {
      return 'Missing required option url';
    }

    if (!args.options.key) {
      return 'Missing required option key';
    }

    return true;
  }
}

module.exports = new SpoPropertyBagRemoveCommand();