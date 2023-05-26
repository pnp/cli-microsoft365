import { IdentitySet, Permission } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { spo } from '../../../../utils/spo';
import { validation } from '../../../../utils/validation';
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

  private getFilteredPermissions(args: CommandArgs, permissions: Permission[]): Permission[] {
    let filterProperty: string = 'displayName';
    let filterValue: string = args.options.appDisplayName as string;

    if (args.options.appId) {
      filterProperty = 'id';
      filterValue = args.options.appId;
    }

    return permissions.filter((p: Permission) => {
      return p.grantedToIdentities!.some(({ application }: IdentitySet) =>
        (application as any)[filterProperty] === filterValue);
    });
  }

  private getApplicationPermission(permissionId: string): Promise<Permission> {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/sites/${this.siteId}/permissions/${permissionId}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.get<Permission>(requestOptions);
  }

  private getTransposed(permissions: Permission[]): { appDisplayName: string; appId: string; permissionId: string, roles: string[] }[] {
    const transposed: { appDisplayName: string; appId: string; permissionId: string, roles: string[] }[] = [];

    permissions.forEach((permissionObject: Permission) => {
      permissionObject.grantedToIdentities!.forEach((permissionEntity: IdentitySet) => {
        transposed.push(
          {
            appDisplayName: permissionEntity.application!.displayName!,
            appId: permissionEntity.application!.id!,
            permissionId: permissionObject.id!,
            roles: permissionObject.roles!
          });
      });
    });

    return transposed;
  }

  private getPermissions(): Promise<{ value: Permission[] }> {
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
      this.siteId = await spo.getSpoGraphSiteId(args.options.siteUrl);
      const permRes: { value: Permission[] } = await this.getPermissions();
      let permissions: Permission[] = permRes.value;

      if (args.options.appId || args.options.appDisplayName) {
        permissions = this.getFilteredPermissions(args, permRes.value);
      }

      const res: Permission[] = await Promise.all(permissions.map(g => this.getApplicationPermission(g.id!)));
      logger.log(this.getTransposed(res));

    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoSiteAppPermissionListCommand();