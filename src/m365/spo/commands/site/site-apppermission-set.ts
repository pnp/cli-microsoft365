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
  appId?: string;
  appDisplayName?: string;
  id?: string;
  siteUrl: string;
}

class SpoSiteAppPermissionSetCommand extends GraphCommand {
  private siteId: string = '';
  private roles: string[] = ['read', 'write', 'manage', 'fullcontrol'];

  public get name(): string {
    return commands.SITE_APPPERMISSION_SET;
  }

  public get description(): string {
    return 'Updates a specific application permission for a site';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: typeof args.options.id !== 'undefined',
        appId: typeof args.options.appId !== 'undefined',
        appDisplayName: typeof args.options.appDisplayName !== 'undefined',
        permissions: args.options.permissions
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --siteUrl <siteUrl>'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '--appId [appId]'
      },
      {
        option: '-n, --appDisplayName [appDisplayName]'
      },
      {
        option: '-p, --permission <permission>',
        autocomplete: this.roles
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.appId && !validation.isValidGuid(args.options.appId)) {
          return `${args.options.appId} is not a valid GUID`;
        }

        if (this.roles.indexOf(args.options.permission) === -1) {
          return `${args.options.permission} is not a valid permission value. Allowed values are ${this.roles.join('|')}`;
        }

        return validation.isValidSharePointUrl(args.options.siteUrl);
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'appId', 'appDisplayName'] });
  }

  private getFilteredPermissions(args: CommandArgs, permissions: Permission[]): Permission[] {
    let filterProperty: string = 'displayName';
    let filterValue: string = args.options.appDisplayName as string;

    if (args.options.appId) {
      filterProperty = 'id';
      filterValue = args.options.appId;
    }

    return permissions.filter((p: Permission) =>
      p.grantedToIdentities!.some(({ application }: IdentitySet) =>
        (application as any)[filterProperty] === filterValue)
    );
  }

  private async getPermission(args: CommandArgs): Promise<string> {
    if (args.options.id) {
      return args.options.id;
    }

    const permissionRequestOptions: any = {
      url: `${this.resource}/v1.0/sites/${this.siteId}/permissions`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response: { value: Permission[] } = await request.get<{ value: Permission[] }>(permissionRequestOptions);

    const sitePermissionItems: Permission[] = this.getFilteredPermissions(args, response.value);

    if (sitePermissionItems.length === 0) {
      throw 'The specified app permission does not exist';
    }

    if (sitePermissionItems.length > 1) {
      throw `Multiple app permissions with displayName ${args.options.appDisplayName} found: ${response.value.map(x => x.grantedToIdentities!.map(y => y.application!.id))}`;
    }

    return sitePermissionItems[0].id!;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      this.siteId = await spo.getSpoGraphSiteId(args.options.siteUrl);
      const sitePermissionId: string = await this.getPermission(args);
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

      const res = await request.patch(requestOptions);
      logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoSiteAppPermissionSetCommand();