import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import GraphCommand from '../../../base/GraphCommand';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface AppInfo {
  appId: string;
  displayName: string;
}

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteUrl: string;
  permission: string;
  appId?: string;
  appDisplayName?: string;
}

class SpoSiteAppPermissionAddCommand extends GraphCommand {
  private siteId: string = '';

  public get name(): string {
    return commands.SITE_APPPERMISSION_ADD;
  }

  public get description(): string {
    return 'Adds an application permissions to the site';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.appId = typeof args.options.appId !== 'undefined';
    telemetryProps.appDisplayName = typeof args.options.appDisplayName !== 'undefined';
    return telemetryProps;
  }

  private getSpoSiteId(args: CommandArgs): Promise<string> {
    const url = new URL(args.options.siteUrl);
    const requestOptions: any = {
      url: `${this.resource}/v1.0/sites/${url.hostname}:${url.pathname}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ id: string }>(requestOptions)
      .then((site: { id: string }) => site.id);
  }

  private getAppInfo(args: CommandArgs): Promise<AppInfo> {
    if (args.options.appId && args.options.appDisplayName) {
      return Promise.resolve({
        appId: args.options.appId as string,
        displayName: args.options.appDisplayName as string
      });
    }

    let endpoint: string = "";

    if (args.options.appId) {
      endpoint = `${this.resource}/v1.0/myorganization/applications?$filter=appId eq '${encodeURIComponent(args.options.appId as string)}'`;
    }
    else {
      endpoint = `${this.resource}/v1.0/myorganization/applications?$filter=displayName eq '${encodeURIComponent(args.options.appDisplayName as string)}'`;
    }

    const appRequestOptions: any = {
      url: endpoint,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: AppInfo[] }>(appRequestOptions)
      .then(response => {
        const appItem: AppInfo | undefined = response.value[0];

        if (!appItem) {
          return Promise.reject("The specified Azure AD app does not exist");
        }

        if (response.value.length > 1) {
          return Promise.reject(`Multiple Azure AD app with displayName ${args.options.appDisplayName} found: ${response.value.map(x => x.appId)}`);
        }

        return Promise.resolve({
          appId: appItem.appId,
          displayName: appItem.displayName
        });
      });
  }

  private mapRequestBody(permission: string, appInfo: AppInfo): any {
    const requestBody: any = {
      roles: permission.split(',')
    };

    requestBody.grantedToIdentities = [];
    requestBody.grantedToIdentities.push({ application: { "id": appInfo.appId, "displayName": appInfo.displayName } });

    return requestBody;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getSpoSiteId(args)
      .then((siteId: string): Promise<AppInfo> => {
        this.siteId = siteId;
        return this.getAppInfo(args);
      })
      .then((appInfo: AppInfo): Promise<any> => {
        const requestBody: any = this.mapRequestBody(args.options.permission, appInfo);

        const requestOptions: any = {
          url: `${this.resource}/v1.0/sites/${this.siteId}/permissions`,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'content-type': 'application/json;odata=nometadata'
          },
          data: requestBody,
          responseType: 'json'
        };

        return request.post(requestOptions);
      })
      .then((res: any): void => {
        logger.log(res);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --siteUrl <siteUrl>'
      },
      {
        option: '-p, --permission <permission>'
      },
      {
        option: '-i, --appId [appId]'
      },
      {
        option: '-n, --appDisplayName [appDisplayName]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.appId && !args.options.appDisplayName) {
      return `Specify appId or appDisplayName, one is required`;
    }

    if (args.options.appId && !Utils.isValidGuid(args.options.appId)) {
      return `${args.options.appId} is not a valid GUID`;
    }

    const permissions: string[] = args.options.permission.split(',');
    for (let i = 0; i < permissions.length; i++) {
      if (['read', 'write'].indexOf(permissions[i]) === -1) {
        return `${permissions[i]} is not a valid permission value. Allowed values read|write|read,write`;
      }
    }

    return SpoCommand.isValidSharePointUrl(args.options.siteUrl);
  }
}

module.exports = new SpoSiteAppPermissionAddCommand();
