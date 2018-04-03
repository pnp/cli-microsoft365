import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import * as request from 'request-promise-native';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import { Auth } from '../../../../Auth';
import { SpoPropertyBagBaseCommand, IdentityResponse } from './propertybag-base';
import GlobalOptions from '../../../../GlobalOptions';
import Utils from '../../../../Utils';

const vorpal: Vorpal = require('../../../../vorpal-init');

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
    return `${commands.PROPERTYBAG_REMOVE}`;
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

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const removeProperty = (): void => {
      const resource: string = Auth.getResourceFromUrl(args.options.webUrl);

      if (this.debug) {
        cmd.log(`Retrieving access token for ${resource}...`);
      }

      auth
        .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
        .then((accessToken: string): request.RequestPromise => {
          this.siteAccessToken = accessToken;

          if (this.debug) {
            cmd.log(`Retrieved access token ${accessToken}. Retrieving request digest...`);
          }

          return this.getRequestDigestForSite(args.options.webUrl, this.siteAccessToken, cmd, this.debug);
        })
        .then((contextResponse: ContextInfo): Promise<IdentityResponse> => {
          this.formDigestValue = contextResponse.FormDigestValue;

          if (this.debug) {
            cmd.log('Response:');
            cmd.log(JSON.stringify(contextResponse));
            cmd.log('');
          }

          return this.requestObjectIdentity(args.options.webUrl, cmd);
        })
        .then((identityResp: IdentityResponse): Promise<IdentityResponse> => {
          if (this.debug) {
            cmd.log('Response:');
            cmd.log(JSON.stringify(identityResp));
            cmd.log('');
          }

          const opts: Options = args.options;
          if (opts.folder) {
            // get the folder guid instead of the web guid
            return this.requestFolderObjectIdentity(identityResp, opts.webUrl, opts.folder, cmd)
          }
          return new Promise<IdentityResponse>(resolve => { return resolve(identityResp); });
        })
        .then((identityResp: IdentityResponse): Promise<any> => {
          if (this.debug) {
            cmd.log('Response:');
            cmd.log(JSON.stringify(identityResp));
            cmd.log('');
          }

          return this.removeProperty(identityResp, args.options);
        })
        .then((res: any): void => {
          if (this.debug) {
            cmd.log('Response:');
            cmd.log(JSON.stringify(res));
            cmd.log('');
          }

          if (this.verbose) {
            cmd.log('DONE');
          }

          cb();
        }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
    }

    if (args.options.confirm) {
      removeProperty();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the ${args.options.key} property?`,
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
      headers: Utils.getRequestHeaders({
        authorization: `Bearer ${this.siteAccessToken}`,
        'X-RequestDigest': this.formDigestValue
      }),
      body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetFieldValue" Id="206" ObjectPathId="205"><Parameters><Parameter Type="String">${Utils.escapeXml(options.key)}</Parameter><Parameter Type="Null" /></Parameters></Method><Method Name="Update" Id="207" ObjectPathId="198" /></Actions><ObjectPaths><Property Id="205" ParentId="198" Name="${objectType}" /><Identity Id="198" Name="${identityResp.objectIdentity}" /></ObjectPaths></Request>`
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
      }, (err: any): void => { reject(err); })
    });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'The URL of the site from which the property should be removed'
      },
      {
        option: '-k, --key <key>',
        description: 'Key of the property to be removed. Case-sensitive'
      },
      {
        option: '-f, --folder [folder]',
        description: 'Site-relative URL of the folder from which to remove the property bag value',
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removal of property bag value'
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

      return true;
    };
  }

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.PROPERTYBAG_REMOVE).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site,
    using the ${chalk.blue(commands.CONNECT)} command.
                      
  Remarks:

    To remove property bag value, you have to first connect to a SharePoint
    Online site using the ${chalk.blue(commands.CONNECT)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.

  Examples:

    Removes the value of the ${chalk.grey('key1')} property from the property bag located in site ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${chalk.grey(config.delimiter)} ${commands.PROPERTYBAG_REMOVE} --webUrl https://contoso.sharepoint.com/sites/test --key key1

    Removes the value of the ${chalk.grey('key1')} property from the property bag located in site root folder ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${chalk.grey(config.delimiter)} ${commands.PROPERTYBAG_REMOVE} --webUrl https://contoso.sharepoint.com/sites/test --key key1 --folder / --confirm

    Removes the value of the ${chalk.grey('key1')} property from the property bag located in site document library ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${chalk.grey(config.delimiter)} ${commands.PROPERTYBAG_REMOVE} --webUrl https://contoso.sharepoint.com/sites/test --key key1 --folder '/Shared Documents'
    
    Removes property bag value located in folder in site document library ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${chalk.grey(config.delimiter)} ${commands.PROPERTYBAG_REMOVE} --webUrl https://contoso.sharepoint.com/sites/test --key key1 --folder '/Shared Documents/MyFolder'

    Removes the value of the ${chalk.grey('key1')} property from the property bag located in site list ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${chalk.grey(config.delimiter)} ${commands.PROPERTYBAG_REMOVE} --webUrl https://contoso.sharepoint.com/sites/test --key key1 --folder /Lists/MyList
    `);
  }
}

module.exports = new SpoPropertyBagRemoveCommand();