import { Logger } from '../../../../cli';
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
  siteUrl: string;
  appId?: string;
  appDisplayName?: string;
}

class SpoSiteAppPermissionListCommand extends GraphCommand {
  private siteId: string = '';

  public get name(): string {
    return commands.SITE_APPPERMISSION_LIST;
  }

  public get description(): string {
    return 'Lists application permissions for a site';
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
        appDisplayName: typeof args.options.appDisplayName !== 'undefined',
        appId: typeof args.options.appId !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --siteUrl <siteUrl>'
      },
      {
        option: '-i, --appId [appId]'
      },
      {
        option: '-n, --appDisplayName [appDisplayName]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.appId && args.options.appDisplayName) {
          return `Provide either appId or appDisplayName, not both`;
        }

        return validation.isValidSharePointUrl(args.options.siteUrl);
      }
    );
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

  private getFilteredPermissions(args: CommandArgs, permissions: SitePermission[]): SitePermission[] {
    let filterProperty: string = 'displayName';
    let filterValue: string = args.options.appDisplayName as string;

    if (args.options.appId) {
      filterProperty = 'id';
      filterValue = args.options.appId;
    }

    return permissions.filter((p: SitePermission) => {
      return p.grantedToIdentities.some(({ application }: SitePermissionIdentitySet) =>
        (application as any)[filterProperty] === filterValue);
    });
  }

  private getApplicationPermission(permissionId: string): Promise<SitePermission> {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/sites/${this.siteId}/permissions/${permissionId}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.get<SitePermission>(requestOptions);
  }

  private getTransposed(permissions: SitePermission[]): { appDisplayName: string; appId: string; permissionId: string, roles: string[] }[] {
    const transposed: { appDisplayName: string; appId: string; permissionId: string, roles: string[] }[] = [];

    permissions.forEach((permissionObject: SitePermission) => {
      permissionObject.grantedToIdentities.forEach((permissionEntity: SitePermissionIdentitySet) => {
        transposed.push(
          {
            appDisplayName: permissionEntity.application.displayName,
            appId: permissionEntity.application.id,
            permissionId: permissionObject.id,
            roles: permissionObject.roles
          });
      });
    });

    return transposed;
  }

  private getPermissions(): Promise<{ value: SitePermission[] }> {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/sites/${this.siteId}/permissions`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.get(requestOptions);
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      this.siteId = await this.getSpoSiteId(args);
      const permRes: { value: SitePermission[] } = await this.getPermissions();
      let permissions: SitePermission[] = permRes.value;

      if (args.options.appId || args.options.appDisplayName) {
        permissions = this.getFilteredPermissions(args, permRes.value);
      }

      const res: SitePermission[] = await Promise.all(permissions.map(g => this.getApplicationPermission(g.id)));
      logger.log(this.getTransposed(res));

    } 
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoSiteAppPermissionListCommand();