import Command, { CommandAction } from '../../Command';
import appInsights from '../../appInsights';
import auth from './SpoAuth';
import { ContextInfo } from './spo';
import * as request from 'request-promise-native';

export default abstract class SpoCommand extends Command {
  protected requiresTenantAdmin(): boolean {
    return false;
  }

  public action(): CommandAction {
    const cmd: SpoCommand = this;

    return function (this: CommandInstance, args: any, cb: () => void) {
      cmd._verbose = args.options.verbose || false;

      appInsights.trackEvent({
        name: cmd.getCommandName(),
        properties: cmd.getTelemetryProperties(args)
      });

      if (!auth.site.connected) {
        this.log('Connect to a SharePoint Online site first');
        cb();
        return;
      }

      if (cmd.requiresTenantAdmin()) {
        if (!auth.site.isTenantAdminSite()) {
          this.log(`${auth.site.url} is not a tenant admin site. Connect to your tenant admin site and try again`);
          cb();
          return;
        }
      }

      cmd.commandAction(this, args, cb);
    }
  }

  protected getRequestDigest(cmd: CommandInstance, verbose: boolean = false): Promise<ContextInfo> {
    const requestOptions: any = {
      url: `${auth.site.url}/_api/contextinfo`,
      headers: {
        authorization: `Bearer ${auth.site.accessToken}`,
        accept: 'application/json;odata=nometadata'
      },
      json: true
    };

    if (verbose) {
      cmd.log('Executing web request...');
      cmd.log(requestOptions);
      cmd.log('');
    }

    return request.post(requestOptions);
  }

  public static isValidSharePointUrl(url: string): boolean | string {
    if (!url) {
      return false;
    }

    if (url.indexOf('https://') !== 0 ||
      url.indexOf('.sharepoint.com') === -1) {
      return `${url} is not a valid SharePoint Online site URL`;
    }
    else {
      return true;
    }
  }
}