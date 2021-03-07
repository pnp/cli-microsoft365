import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { SitePermission, SitePermissionIdentitySet } from './SitePermission';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteUrl: string;
  permissionId: string
}


class SpoSiteApppermissionGetCommand extends GraphCommand {
  public get name(): string {
    return commands.SITE_APPPERMISSION_GET;
  }

  public get description(): string {
    return 'Get a specific application permissions for the site';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.permissionId = typeof args.options.permissionId !== 'undefined';
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

  private getApplicationPermission(args: CommandArgs, siteId: string): Promise<SitePermission> {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/sites/${siteId}/permissions/${args.options.permissionId}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.get(requestOptions);
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    this
      .getSpoSiteId(args)
      .then((siteId: string): Promise<SitePermission> => this.getApplicationPermission(args, siteId))
      .then((permissionObject: SitePermission) => {
        const transposed: { appDisplayName: string; appId: string; permissionId: string }[] = [];

        permissionObject.grantedToIdentities.forEach((permissionEntity: SitePermissionIdentitySet) => {
          transposed.push(
            {
              appDisplayName: permissionEntity.application.displayName,
              appId: permissionEntity.application.id,
              permissionId: permissionObject.id
            });
        });

        logger.log(transposed);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --siteUrl <siteUrl>'
      },
      {
        option: '-i, --permissionId <permissionId>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    return SpoCommand.isValidSharePointUrl(args.options.siteUrl);
  }
}

module.exports = new SpoSiteApppermissionGetCommand();
