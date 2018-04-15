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
import { BasePermissions, PermissionKind } from '../../common/base-permissions';

const vorpal: Vorpal = require('../../../../vorpal-init');

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
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    let webIdentityResp: IdentityResponse;

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
      .then((identityResp: IdentityResponse): Promise<boolean> => {
        webIdentityResp = identityResp;

        // Check if web no script enabled or not
        // Cannot set property bag value if no script is enabled
        return this.isNoScriptSite(identityResp, args.options, cmd);
      })
      .then((isNoScriptSite: boolean): Promise<IdentityResponse> => {
        if (isNoScriptSite) {
          return Promise.reject('Site has NoScript enabled, and setting property bag values is not supported');
        }

        const opts: Options = args.options;
        if (opts.folder) {
          // get the folder guid instead of the web guid
          return this.requestFolderObjectIdentity(webIdentityResp, opts.webUrl, opts.folder, cmd);
        }

        return new Promise<IdentityResponse>(resolve => { return resolve(webIdentityResp); });
      })
      .then((identityResp: IdentityResponse): Promise<any> => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(JSON.stringify(identityResp));
          cmd.log('');
        }

        return this.setProperty(identityResp, args.options);
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

  private setProperty(identityResp: IdentityResponse, options: Options): Promise<any> {
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
      body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetFieldValue" Id="206" ObjectPathId="205"><Parameters><Parameter Type="String">${Utils.escapeXml(options.key)}</Parameter><Parameter Type="String">${Utils.escapeXml(options.value)}</Parameter></Parameters></Method><Method Name="Update" Id="207" ObjectPathId="198" /></Actions><ObjectPaths><Property Id="205" ParentId="198" Name="${objectType}" /><Identity Id="198" Name="${identityResp.objectIdentity}" /></ObjectPaths></Request>`
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

  /**
   * Detects if the site in question has no script enabled or not. Detection is done
   * by verifying if the AddAndCustomizePages permission is missing
   * Note: Can later be moved as common method if required for other cli checks.
   * @param webIdentityResp web object identity response returned from client.svc/ProcessQuery. Has format like <GUID>|<GUID>:site:<GUID>:web:<GUID>
   * @param options command options
   * @param cmd command instance
   */
  private isNoScriptSite(webIdentityResp: IdentityResponse, options: Options, cmd: CommandInstance): Promise<boolean> {
    return new Promise<boolean>((resolve: (isNoScriptSite: boolean) => void, reject: (error: any) => void): void => {
      this.getEffectiveBasePermissions(webIdentityResp.objectIdentity, options.webUrl, cmd)
        .then((basePermissionsResp: BasePermissions): void => {
          resolve(basePermissionsResp.has(PermissionKind.AddAndCustomizePages) === false);
        })
        .catch(err => reject(err));
    });
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

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.PROPERTYBAG_SET).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site,
    using the ${chalk.blue(commands.CONNECT)} command.
                       
  Remarks:

    To set property bag value, you have to first connect to a SharePoint
    Online site using the ${chalk.blue(commands.CONNECT)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.

  Examples:

    Sets the value of the ${chalk.grey('key1')} property in the property bag of site
    ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${chalk.grey(config.delimiter)} ${commands.PROPERTYBAG_SET} --webUrl https://contoso.sharepoint.com/sites/test --key key1 --value value1

    Sets the value of the ${chalk.grey('key1')} property in the property bag of the root folder
    of site ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${chalk.grey(config.delimiter)} ${commands.PROPERTYBAG_SET} --webUrl https://contoso.sharepoint.com/sites/test --key key1 --value value1 --folder /

    Sets the value of the ${chalk.grey('key1')} property in the property bag of a document
    library located in site ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${chalk.grey(config.delimiter)} ${commands.PROPERTYBAG_SET} --webUrl https://contoso.sharepoint.com/sites/test --key key1 --value value1 --folder '/Shared Documents'
    
    Sets the value of the ${chalk.grey('key1')} property in the property bag of a folder
    in a document library located in site
    ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${chalk.grey(config.delimiter)} ${commands.PROPERTYBAG_SET} --webUrl https://contoso.sharepoint.com/sites/test --key key1 --value value1 --folder '/Shared Documents/MyFolder'

    Sets the value of the ${chalk.grey('key1')} property in the property bag of a list in site
    ${chalk.grey('https://contoso.sharepoint.com/sites/test')}
      ${chalk.grey(config.delimiter)} ${commands.PROPERTYBAG_SET} --webUrl https://contoso.sharepoint.com/sites/test --key key1 --value value1 --folder /Lists/MyList
    `);
  }
}

module.exports = new SpoPropertyBagSetCommand();