import { Logger } from '../../../../cli';
import {
  CommandError
} from '../../../../Command';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { ClientSvcResponse, ClientSvcResponseContents, ContextInfo, spo, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import { BasePermissions, PermissionKind } from '../../base-permissions';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  title: string;
  webUrl: string;
  webTemplate: string;
  parentWebUrl: string;
  description?: string;
  locale?: string;
  breakInheritance: boolean;
  inheritNavigation: boolean;
}

class SpoWebAddCommand extends SpoCommand {
  // used to early break promises chain
  private static DONE: string = 'DONE';

  public get name(): string {
    return commands.WEB_ADD;
  }

  public get description(): string {
    return 'Create new subsite';
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
        description: (!(!args.options.description)).toString(),
        locale: args.options.locale || '1033',
        breakInheritance: args.options.breakInheritance || false,
        inheritNavigation: args.options.inheritNavigation || false
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-t, --title <title>'
      },
      {
        option: '-d, --description [description]'
      },
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-w, --webTemplate <webTemplate>'
      },
      {
        option: '-p, --parentWebUrl <parentWebUrl>'
      },
      {
        option: '-l, --locale [locale]'
      },
      {
        option: '--breakInheritance'
      },
      {
        option: '--inheritNavigation'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.parentWebUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.locale) {
          const locale: number = parseInt(args.options.locale);
          if (isNaN(locale)) {
            return `${args.options.locale} is not a valid locale number`;
          }
        }

        return true;
      }
    );
  }

  protected getExcludedOptionsWithUrls(): string[] | undefined {
    return ['webUrl'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    let siteInfo: any = null;
    let subsiteFullUrl: string = '';

    spo
      .getRequestDigest(args.options.parentWebUrl)
      .then((res: ContextInfo): Promise<any> => {
        const requestOptions: any = {
          url: `${args.options.parentWebUrl}/_api/web/webinfos/add`,
          headers: {
            'content-type': 'application/json;odata=nometadata',
            accept: 'application/json;odata=nometadata',
            'X-RequestDigest': res.FormDigestValue
          },
          responseType: 'json',
          data: {
            parameters: {
              Url: args.options.webUrl,
              Title: args.options.title,
              Description: args.options.description,
              Language: args.options.locale,
              WebTemplate: args.options.webTemplate,
              UseUniquePermissions: args.options.breakInheritance
            }
          }
        };

        if (this.verbose) {
          logger.logToStderr(`Creating subsite ${args.options.parentWebUrl}/${args.options.webUrl}...`);
        }

        return request.post(requestOptions);
      })
      .then((res: any): Promise<any> => {
        siteInfo = res;

        if (!args.options.inheritNavigation) {
          return Promise.reject(SpoWebAddCommand.DONE);
        }

        if (this.verbose) {
          logger.logToStderr("Setting inheriting navigation from the parent site...");
        }

        subsiteFullUrl = `${args.options.parentWebUrl}/${encodeURIComponent(args.options.webUrl)}`;

        const requestOptions: any = {
          url: `${subsiteFullUrl}/_api/web/effectivebasepermissions`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        return request.get(requestOptions);
      })
      .then((res: any): Promise<ContextInfo> => {
        const permissions: BasePermissions = new BasePermissions();
        permissions.high = res.High as number;
        permissions.low = res.Low as number;

        /// Detects if the site in question has no script enabled or not. 
        /// Detection is done by verifying if the AddAndCustomizePages permission is missing.
        /// 
        /// See https://support.office.com/en-us/article/Turn-scripting-capabilities-on-or-off-1f2c515f-5d7e-448a-9fd7-835da935584f
        /// for the effects of NoScript
        if (!permissions.has(PermissionKind.AddAndCustomizePages)) {
          if (this.verbose) {
            logger.logToStderr("No script is enabled. Skipping the InheritParentNavigation settings.");
          }

          return Promise.reject(SpoWebAddCommand.DONE);
        }

        return spo.getRequestDigest(subsiteFullUrl);
      })
      .then((res: ContextInfo): Promise<string> => {
        const requestOptions: any = {
          url: `${subsiteFullUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': res.FormDigestValue
          },
          data: `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}"><Actions><ObjectPath Id="1" ObjectPathId="0" /><ObjectPath Id="3" ObjectPathId="2" /><ObjectPath Id="5" ObjectPathId="4" /><SetProperty Id="6" ObjectPathId="4" Name="UseShared"><Parameter Type="Boolean">true</Parameter></SetProperty></Actions><ObjectPaths><StaticProperty Id="0" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /><Property Id="2" ParentId="0" Name="Web" /><Property Id="4" ParentId="2" Name="Navigation" /></ObjectPaths></Request>`
        };

        return request.post(requestOptions);
      })
      .then((res: string): void => {
        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          cb(new CommandError(response.ErrorInfo.ErrorMessage));
          return;
        }
        else {
          logger.log(siteInfo);
        }

        cb();
      }, (err: any): void => {
        if (err === SpoWebAddCommand.DONE) {
          logger.log(siteInfo);


          cb();
          return;
        }

        if (err.error &&
          err.error['odata.error'] &&
          err.error['odata.error'].message) {
          cb(new CommandError(err.error['odata.error'].message.value));
        }
        else {
          this.handleRejectedPromise(err, logger, cb);
        }
      });
  }
}

module.exports = new SpoWebAddCommand();