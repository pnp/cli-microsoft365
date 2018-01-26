import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import * as request from 'request-promise-native';
import GlobalOptions from '../../../../GlobalOptions';
import { ContextInfo } from '../../spo';
import {
  CommandOption,
  CommandValidate,
  CommandError,
  
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Auth from '../../../../Auth';
import Utils from '../../../../Utils';
import { PermissionKind, BasePermissions } from '../customaction/base-permissions';
const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  title?: string;
  webUrl: string;
  webTemplate?: string;
  parentWebUrl: string;
  description?: string;
  locale?: string;
  breakInheritance: boolean;
  inheritNavigation: boolean;
}

class WebAddCommand extends SpoCommand {

  private createWeb(siteAccessToken:string, cmd: CommandInstance, args: CommandArgs, cb: () => void, debug : boolean) : Promise<any> {
    return this.getRequestDigestForSite(args.options.parentWebUrl, siteAccessToken, cmd, this.debug).then((res: ContextInfo): Promise<any> => {
    let requestOptions: any = {
      url: `${args.options.parentWebUrl}/_api/web/webinfos/add`,
      headers: Utils.getRequestHeaders({
        authorization: `Bearer ${siteAccessToken}`,
        'content-type': 'application/json;odata=nometadata',
        accept: 'application/json;odata=nometadata',
        "X-RequestDigest": res.FormDigestValue
      }),
      json: true,
      body: {
        'parameters': {
          'Url': args.options.webUrl,
          'Title': args.options.title,
          'Description': args.options.description,
          'Language':args.options.locale,
          'WebTemplate':args.options.webTemplate,
          'UseUniquePermissions':args.options.breakInheritance,
        }
      }};

      return request.post(requestOptions).then((res:any) : any => {
        cmd.log(`Subsite ${args.options.title} created.`)
        cmd.log(res);
        return res;
      }, (err: any) =>  { 
        cmd.log(new CommandError(`Failed to create the web - ${args.options.webUrl}`)); 
        return Promise.reject(err);
      });
    });
  }

  private getEffectiveBasePermission(siteAccessToken:string, cmd: CommandInstance, args: CommandArgs, cb: () => void, debug : boolean) : Promise<BasePermissions> {
    let subsiteFullUrl = `${args.options.parentWebUrl}/${args.options.webUrl}`;
    
    return this.getRequestDigestForSite(subsiteFullUrl, siteAccessToken, cmd, debug)
    .then((res: ContextInfo): Promise<any> => {
      let requestOptions: any = {
        url: `${subsiteFullUrl}/_api/web/effectivebasepermissions`,
        headers: Utils.getRequestHeaders({
          authorization: `Bearer ${siteAccessToken}`,
          'content-type': 'application/json;odata=nometadata',
          accept: 'application/json;odata=nometadata',
          "X-RequestDigest": res.FormDigestValue
        }),
        json: true
      };

      return request.get(requestOptions).then((res:any) : any => {
        let webEffectivePermission : BasePermissions = new BasePermissions();
        webEffectivePermission.high = res.High as number;
        webEffectivePermission.low = res.Low as number;
        if (debug) {
          cmd.log("Response : WebEffectiveBasePermission")
          cmd.log(res);
        }
  
        return webEffectivePermission;
      }, (err: any) =>  { 
        cmd.log(new CommandError(`Failed to get the effectivebasepermission for the web - ${subsiteFullUrl}`)); 
        return Promise.reject(err);
      });
    });
  }

  private setInheritNavigation(siteAccessToken:string, cmd: CommandInstance, args: CommandArgs, cb: () => void, debug : boolean) : Promise<boolean> {
    let subsiteFullUrl = `${args.options.parentWebUrl}/${args.options.webUrl}`;
  
    return this.getRequestDigestForSite(subsiteFullUrl, siteAccessToken, cmd, debug)
    .then((res: ContextInfo): Promise<any> => {
      let requestOptions: any = {
        url: `${subsiteFullUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: Utils.getRequestHeaders({
          authorization: `Bearer ${siteAccessToken}`,
          'X-RequestDigest': res.FormDigestValue
        }),
        body: `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="Javascript Library"><Actions><ObjectPath Id="1" ObjectPathId="0" /><ObjectPath Id="3" ObjectPathId="2" /><ObjectPath Id="5" ObjectPathId="4" /><SetProperty Id="6" ObjectPathId="4" Name="UseShared"><Parameter Type="Boolean">true</Parameter></SetProperty><Query Id="7" ObjectPathId="4"><Query SelectAllProperties="true"><Properties><Property Name="UseShared" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><StaticProperty Id="0" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /><Property Id="2" ParentId="0" Name="Web" /><Property Id="4" ParentId="2" Name="Navigation" /></ObjectPaths></Request>`
      };

      return request.post(requestOptions).then((res:any) : any => {
        if (debug) {
          cmd.log("Response : SetInheritNavigation");
          cmd.log(res);
        }
        return res;
      }, (err: any) =>  { 
        cmd.log(new CommandError(`Failed to set inheritNavigation for the web - ${subsiteFullUrl}`)); 
        return Promise.reject(err);
      });
    });
  }

  public get name(): string {
    return commands.WEB_ADD;
  }

  public get description(): string {
    return 'Creates new subsite';
  }

  protected requiresTenantAdmin(): boolean {
    return false;
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.title = (!(!args.options.title)).toString();
    telemetryProps.webUrl = (!(!args.options.webUrl)).toString();
    telemetryProps.webTemplate = (!(!args.options.webTemplate)).toString();
    telemetryProps.parentWebUrl = (!(!args.options.parentWebUrl)).toString();
    telemetryProps.description = (!(!args.options.description)).toString();
    telemetryProps.locale = (!(!args.options.locale)).toString();
    telemetryProps.breakInheritance = args.options.breakInheritance || false;
    telemetryProps.inheritNavigation = args.options.inheritNavigation || false;

    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.parentWebUrl);
    let siteAccessToken: string = '';

    if (this.debug) {
      cmd.log(`Retrieving access token for ${resource}...`);
    }
    
    auth
    .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
    .then((accessToken: string): Promise<ContextInfo> => {
      if (this.debug) {
        cmd.log(`Retrieved access token ${accessToken}. Retrieving request digest...`);
      }
      return this.createWeb(accessToken, cmd, args, cb, this.debug);
    })
    .then((res: any) : any => {
      if(args.options.inheritNavigation)
      {
        if(this.debug)
        {
          cmd.log("Setting the navigation to inherit the parent settings.");
        }

        this.getEffectiveBasePermission(siteAccessToken, cmd, args, cb, this.debug)
        .then((perm:BasePermissions) : any => {
            /// Detects if the site in question has no script enabled or not. 
            /// Detection is done by verifying if the AddAndCustomizePages permission is missing.
            /// 
            /// See https://support.office.com/en-us/article/Turn-scripting-capabilities-on-or-off-1f2c515f-5d7e-448a-9fd7-835da935584f
            /// for the effects of NoScript
            /// 
            if(perm.has(PermissionKind.AddAndCustomizePages)) {
              cmd.log("Setting the Navigation to inherit the parent site.")
              return this.setInheritNavigation(siteAccessToken, cmd, args, cb, this.debug).then((res : any) => {
                cb();
              },(reason: any) =>  { cmd.log(reason); cb();})
            }
            else {
              cmd.log("No script is enabled. Skipping the InheitParentNvaigation settings.")
              cb();
            }
          }, (reason: any) =>  { cmd.log(reason); cb();});
      }
      else
       {
         cb();
       }
    }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-t, --title <title>',
        description: 'Subsite title'
      },
      {
        option: '-d, --description [description]',
        description: 'Subsite description, optional'
      },
      {
        option: '--webUrl [webUrl]',
        description: 'Subsite web relative url'
      },
      {
        option: '--webTemplate [webTemplate]',
        description: 'Subsite template, eg. STS#0 (Classic team site)'
      },
      {
        option: '--parentWebUrl [parentWebUrl]',
        description: 'URL of the parent site under which to create the subsite'
      },
      {
        option: '--locale [locale]',
        description: 'Subsite locale LCID, eg. 1033 for en-US'
      },
      {
        option: '--breakInheritance [breakInheritance]',
        description: 'Set to not inherit permissions from the parent site, optional'
      },
      {
        option: '--inheritNavigation [inheritNavigation]',
        description: 'Set to inherit the navigation from the parent site, optional'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
     
      if (!args.options.title) {
        return 'Required option title missing';
      }

      if (!args.options.webUrl) {
        return 'Required option webUrl missing';
      }

      if (!args.options.webTemplate) {
        return 'Required option webTemplate missing';
      }

      if (!args.options.parentWebUrl) {
        return 'Required option parentWebUrl missing';
      }

      if (!args.options.locale) {
        return 'Required option locale missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site,
      using the ${chalk.blue(commands.CONNECT)} command.

    Remarks:
    
      To create a subsite, you have to first connect to SharePoint using the
      ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.
    
    Examples:
    
      Create subsite
        ${chalk.grey(config.delimiter)} ${commands.WEB_ADD} --title subsite --description subsite 1 --webUrl "subsite" --webTemplate STS#0 --parentWebUrl https://contoso.sharepoint.com --locale 1033

      Create subsite with breaking permission inheritance
      ${chalk.grey(config.delimiter)} ${commands.WEB_ADD} --title subsite --description subsite 1 --webUrl "subsite" --webTemplate STS#0 --parentWebUrl https://contoso.sharepoint.com --locale 1033 --breakInheritance

      Create subsite with inheriting the navigation
      ${chalk.grey(config.delimiter)} ${commands.WEB_ADD} --title subsite --description subsite 1 --webUrl "subsite" --webTemplate STS#0 --parentWebUrl https://contoso.sharepoint.com --locale 1033 --inheritNavigation

      More information
      
      Creating subsite using REST
      https://docs.microsoft.com/en-us/sharepoint/dev/apis/rest/complete-basic-operations-using-sharepoint-rest-endpoints#creating-a-site-with-rest
  ` );
  }
}

module.exports = new WebAddCommand();