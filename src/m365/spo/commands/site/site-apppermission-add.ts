import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import GraphCommand from '../../../base/GraphCommand';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteUrl: string;
  permission: string;
  appId: string;
  appDisplayName: string;
}

class SpoSiteApppermissionAddCommand extends GraphCommand {
  public get name(): string {
    return commands.SITE_APPPERMISSION_ADD;
  }

  public get description(): string {
    return 'Adds an application permissions to the site';
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

  private mapRequestBody(options: Options): any {
    const applicationInfo: any = {};
    if (options.appId) {
      applicationInfo.id = options.appId;
    }

    if (options.appDisplayName) {
      applicationInfo.displayName = options.appDisplayName;
    }

    return `{
      "roles": ${JSON.stringify(options.permission.split(','))},
      "grantedToIdentities": [{
        "application": ${JSON.stringify(applicationInfo)}
      }]
    }`;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getSpoSiteId(args)
      .then((siteId: string): Promise<void> => {
        const requestBody: any = this.mapRequestBody(args.options);

        const requestOptions: any = {
          url: `${this.resource}/v1.0/sites/${siteId}/permissions`,
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
        option: '-i, --appId <appId>'
      },
      {
        option: '-n, --appDisplayName <appDisplayName>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.appId && !Utils.isValidGuid(args.options.appId)) {
      return `${args.options.appId} is not a valid GUID`;
    }

    let permissions: string[] = args.options.permission.split(',');
    for (let i = 0; i < permissions.length; i++) {
      if (['read', 'write'].indexOf(permissions[i]) === -1) {
        return `${permissions[i]} is not a valid permission value. Allowed values read|write|read,write`;
      }
    }

    return SpoCommand.isValidSharePointUrl(args.options.siteUrl);
  }
}

module.exports = new SpoSiteApppermissionAddCommand();
