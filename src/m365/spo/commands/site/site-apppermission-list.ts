import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteUrl: string;
  appId?: string;
  appDisplayName?: string;
}

class SiteApppermissionListCommand extends GraphCommand {
  public get name(): string {
    return commands.SITE_APPPERMISSION_LIST;
  }

  public get description(): string {
    return 'Lists application permissions for a site';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.appDisplayName = typeof args.options.appDisplayName !== 'undefined';
    return telemetryProps;
  }

  public getSpoSiteId(args: CommandArgs): Promise<string> {
    return new Promise<string>((resolve: (id: string) => void, reject: (error: any) => void): void => {
      const url = new URL(args.options.siteUrl);
      const requestOptions: any = {
        url: `${this.resource}/v1.0/sites/${url.hostname}:${url.pathname}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };
      request.get(requestOptions).then((site: any) => {
        resolve(site.id);
      }, (error: any): void => {
        reject(error);
      });
    });
  }

  public getFilteredPermissions(args: CommandArgs, permissions: any): any {
    let filterProperty = "displayName";
    let filterValue = args.options.appDisplayName;
    if (args.options.appId) {
      filterProperty = "id"
      filterValue = args.options.appId;
    }
    return permissions.value.filter((p: any) => {
      return p.grantedToIdentities.some(({ application }: any) =>
        application[filterProperty] === filterValue)
    })
  }

  public getTransposed(permissions: any): any {
    let transpose = new Array();
    permissions.forEach((permissionObject: any) => {
      permissionObject.grantedToIdentities.forEach((permissionEntity: any) => {
        if (permissionEntity.application) {
          transpose.push(
            {
              appDisplayName: permissionEntity.application.displayName,
              appId: permissionEntity.application.id,
              permissionId: permissionObject.id
            })
        }
      });
    });
    return transpose;
  }

  public getPermission(siteId: string): Promise<any> {
    return new Promise<string>((resolve: (id: string) => void, reject: (error: any) => void): void => {
      const requestOptions: any = {
        url: `${this.resource}/v1.0/sites/${siteId}/permissions`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      request.get(requestOptions).then((permissions: any): void => {
        resolve(permissions);
      }, (error: any): void => {
        reject(error);
      })
    })
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    this.getSpoSiteId(args)
      .then((siteId: string): Promise<any> => this.getPermission(siteId))
      .then((permissions: any) => {
        let data: any = new Array();
        if (args.options.appId || args.options.appDisplayName) {
          data = this.getFilteredPermissions(args, permissions);
        } else {
          data = permissions.value;
        }
        if (args.options.output && args.options.output.toLowerCase() === 'json') {
          logger.log(data);
        }
        else {
          data = this.getTransposed(data);
          logger.log(data);
        }
        cb();

      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --siteUrl <siteUrl>'
      },
      {
        option: '--appId [appId]'
      },
      {
        option: '-n, --appDisplayName [appDisplayName]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.appId && args.options.appDisplayName) {
      return `Provide either appId or appDisplayName, not both`
    }
    try {
      new URL(args.options.siteUrl)
    } catch (error) {
      return `${args.options.siteUrl} is not a valid URL`;
    }
    return true;
  }
}

module.exports = new SiteApppermissionListCommand();