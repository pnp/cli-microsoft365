import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { spo } from '../../../../utils/spo';
import { validation } from '../../../../utils/validation';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { SitePermission, SitePermissionIdentitySet } from './SitePermission';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteUrl: string;
  id: string
}

class SpoSiteAppPermissionGetCommand extends GraphCommand {
  public get name(): string {
    return commands.SITE_APPPERMISSION_GET;
  }

  public get description(): string {
    return 'Get a specific application permissions for the site';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --siteUrl <siteUrl>'
      },
      {
        option: '-i, --id <id>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.siteUrl)
    );
  }

  private getApplicationPermission(args: CommandArgs, siteId: string): Promise<SitePermission> {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/sites/${siteId}/permissions/${args.options.id}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.get(requestOptions);
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const siteId: string = await spo.getSpoGraphSiteId(args.options.siteUrl, logger, this.debug);
      const permissionObject: SitePermission = await this.getApplicationPermission(args, siteId);
      const transposed: { appDisplayName: string; appId: string; permissionId: string, roles: string }[] = [];

      permissionObject.grantedToIdentities.forEach((permissionEntity: SitePermissionIdentitySet) => {
        transposed.push(
          {
            appDisplayName: permissionEntity.application.displayName,
            appId: permissionEntity.application.id,
            permissionId: permissionObject.id,
            roles: permissionObject.roles.join()
          });
      });

      logger.log(transposed);

    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoSiteAppPermissionGetCommand();
