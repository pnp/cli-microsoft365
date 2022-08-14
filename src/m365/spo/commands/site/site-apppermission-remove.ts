import { Cli, Logger } from '../../../../cli';
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
  permissionId?: string;
  confirm?: boolean;
}

class SpoSiteAppPermissionRemoveCommand extends GraphCommand {
  private siteId: string = '';

  public get name(): string {
    return commands.SITE_APPPERMISSION_REMOVE;
  }

  public get description(): string {
    return 'Removes an application permission from the site';
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
        appId: typeof args.options.appId !== 'undefined',
        appDisplayName: typeof args.options.appDisplayName !== 'undefined',
        permissionId: typeof args.options.permissionId !== 'undefined',
        confirm: (!!args.options.confirm).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
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
        option: '--confirm'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.appId && !validation.isValidGuid(args.options.appId)) {
          return `${args.options.appId} is not a valid GUID`;
        }

        return validation.isValidSharePointUrl(args.options.siteUrl);
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(['appId', 'appDisplayName', 'permissionId']);
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

  private getPermissionIds(args: CommandArgs): Promise<string[]> {
    if (args.options.permissionId) {
      return Promise.resolve([args.options.permissionId!]);
    }

    return this
      .getPermissions()
      .then((res: { value: SitePermission[] }) => {
        let permissions: SitePermission[] = res.value;

        if (args.options.appId || args.options.appDisplayName) {
          permissions = this.getFilteredPermissions(args, res.value);
        }

        return Promise.resolve(permissions.map(x => x.id));
      });
  }

  private removePermissions(permissionId: string): Promise<void> {
    const spRequestOptions: any = {
      url: `${this.resource}/v1.0/sites/${this.siteId}/permissions/${permissionId}`,
      headers: {
        'accept': 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.delete(spRequestOptions);
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const removeSiteAppPermission: () => void = (): void => {
      this
        .getSpoSiteId(args)
        .then((siteId: string): Promise<string[]> => {
          this.siteId = siteId;
          return this.getPermissionIds(args);
        })
        .then((permissionIdsToRemove: string[]): Promise<void[]> => {
          const tasks: Promise<void>[] = [];

          for (const permissionId of permissionIdsToRemove) {
            tasks.push(this.removePermissions(permissionId));
          }

          return Promise.all(tasks);
        })
        .then((res: any): void => {
          logger.log(res);
          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    };

    if (args.options.confirm) {
      removeSiteAppPermission();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the specified application permission from site ${args.options.siteUrl}?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeSiteAppPermission();
        }
      });
    }
  }
}

module.exports = new SpoSiteAppPermissionRemoveCommand();
