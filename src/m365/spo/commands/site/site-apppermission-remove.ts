import { IdentitySet, Permission } from '@microsoft/microsoft-graph-types';
import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteUrl: string;
  appId?: string;
  appDisplayName?: string;
  id?: string;
  force?: boolean;
}

class SpoSiteAppPermissionRemoveCommand extends GraphCommand {
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
        id: typeof args.options.id !== 'undefined',
        force: (!!args.options.force).toString()
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
        option: '-f, --force'
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
    this.optionSets.push({ options: ['appId', 'appDisplayName', 'id'] });
  }

  private getPermissions(siteId: string): Promise<{ value: Permission[] }> {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/sites/${siteId}/permissions`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.get(requestOptions);
  }

  private getFilteredPermissions(options: Options, permissions: Permission[]): Permission[] {
    let filterProperty: string = 'displayName';
    let filterValue: string = options.appDisplayName as string;

    if (options.appId) {
      filterProperty = 'id';
      filterValue = options.appId;
    }

    return permissions.filter((p: Permission) => {
      return p.grantedToIdentities!.some(({ application }: IdentitySet) =>
        (application as any)[filterProperty] === filterValue);
    });
  }

  private async getPermissionIds(siteId: string, options: Options): Promise<string[]> {
    if (options.id) {
      return Promise.resolve([options.id!]);
    }

    const permissionsObject = await this.getPermissions(siteId);
    let permissions = permissionsObject.value;

    if (options.appId || options.appDisplayName) {
      permissions = this.getFilteredPermissions(options, permissionsObject.value);
    }

    return permissions.map(x => x.id!);
  }

  private removePermissions(siteId: string, permissionId: string): Promise<void> {
    const spRequestOptions: any = {
      url: `${this.resource}/v1.0/sites/${siteId}/permissions/${permissionId}`,
      headers: {
        'accept': 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.delete(spRequestOptions);
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removeSiteAppPermission(logger, args.options);
    }
    else {
      const result = await Cli.promptForConfirmation(`Are you sure you want to remove the specified application permission from site ${args.options.siteUrl}?`);

      if (result) {
        await this.removeSiteAppPermission(logger, args.options);
      }
    }
  }

  private async removeSiteAppPermission(logger: Logger, options: Options): Promise<void> {
    try {
      const siteId = await spo.getSpoGraphSiteId(options.siteUrl);
      const permissionIdsToRemove: string[] = await this.getPermissionIds(siteId, options);
      const tasks: Promise<void>[] = [];

      for (const permissionId of permissionIdsToRemove) {
        tasks.push(this.removePermissions(siteId, permissionId));
      }

      const response = await Promise.all(tasks);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteAppPermissionRemoveCommand();
