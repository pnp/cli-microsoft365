import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { SitePermission, SitePermissionIdentitySet } from './SitePermission';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appId?: string;
  appDisplayName?: string;
  permissionId?: string;
  siteUrl: string;
}

class SpoSiteAppPermissionSetCommand extends GraphCommand {
  private siteId: string = '';

  public get name(): string {
    return commands.SITE_APPPERMISSION_SET;
  }

  public get description(): string {
    return 'Updates a specific application permission for a site';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.permissionId = typeof args.options.permissionId !== 'undefined';
    telemetryProps.appId = typeof args.options.appId !== 'undefined';
    telemetryProps.appDisplayName = typeof args.options.appDisplayName !== 'undefined';
    return telemetryProps;
  }

  private getSpoSiteId(args: CommandArgs): Promise<string> {
    const url = new URL(args.options.siteUrl);
    const siteRequestOptions: any = {
      url: `${this.resource}/v1.0/sites/${url.hostname}:${url.pathname}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ id: string }>(siteRequestOptions)
      .then((site: { id: string }) => site.id);
  }

  private getFilteredPermissions(args: CommandArgs, permissions: SitePermission[]): SitePermission[] {
    let filterProperty: string = 'displayName';
    let filterValue: string = args.options.appDisplayName as string;

    if (args.options.appId) {
      filterProperty = 'id';
      filterValue = args.options.appId;
    }

    return permissions.filter((p: SitePermission) =>
      p.grantedToIdentities.some(({ application }: SitePermissionIdentitySet) =>
        (application as any)[filterProperty] === filterValue)
    );
  }

  private getPermission(args: CommandArgs): Promise<string> {
    if (args.options.permissionId) {
      return Promise.resolve(args.options.permissionId);
    }

    const permissionRequestOptions: any = {
      url: `${this.resource}/v1.0/sites/${this.siteId}/permissions`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: SitePermission[] }>(permissionRequestOptions)
      .then(response => {
        const sitePermissionItems: SitePermission[] = this.getFilteredPermissions(args, response.value);

        if (sitePermissionItems.length === 0) {
          return Promise.reject('The specified app permission does not exist');
        }

        if (sitePermissionItems.length > 1) {
          return Promise.reject(`Multiple app permissions with displayName ${args.options.appDisplayName} found: ${response.value.map(x => x.grantedToIdentities.map(y => y.application.id))}`);
        }

        return Promise.resolve(sitePermissionItems[0].id);
      });
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    this
      .getSpoSiteId(args)
      .then((siteId: string): Promise<string> => {
        this.siteId = siteId;
        return this.getPermission(args);
      })
      .then((sitePermissionId: string) => {
        const requestOptions: any = {
          url: `${this.resource}/v1.0/sites/${this.siteId}/permissions/${sitePermissionId}`,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'content-type': 'application/json;odata=nometadata'
          },
          data: {
            roles: [args.options.permission]
          },
          responseType: 'json'
        };

        return request.patch(requestOptions);
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
        option: '-i, --permissionId [permissionId]'
      },
      {
        option: '--appId [appId]'
      },
      {
        option: '-n, --appDisplayName [appDisplayName]'
      },
      {
        option: '-p, --permission <permission>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.permissionId && !args.options.appId && !args.options.appDisplayName) {
      return `Specify permissionId, appId or appDisplayName, one is required`;
    }

    if (args.options.appId && !validation.isValidGuid(args.options.appId)) {
      return `${args.options.appId} is not a valid GUID`;
    }

    if (['read', 'write', 'owner'].indexOf(args.options.permission) === -1) {
      return `${args.options.permission} is not a valid permission value. Allowed values are read|write|owner`;
    }

    return validation.isValidSharePointUrl(args.options.siteUrl);
  }
}

module.exports = new SpoSiteAppPermissionSetCommand();
