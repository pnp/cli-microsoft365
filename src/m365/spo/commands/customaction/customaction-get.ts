import request from '../../../../request';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { CustomAction } from './customaction';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  url: string;
  scope?: string;
}

class SpoCustomActionGetCommand extends SpoCommand {
  public get name(): string {
    return `${commands.CUSTOMACTION_GET}`;
  }

  public get description(): string {
    return 'Gets details for the specified custom action';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.scope = args.options.scope || 'All';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    ((): Promise<CustomAction> => {
      if (args.options.scope && args.options.scope.toLowerCase() !== "all") {
        return this.getCustomAction(args.options);
      }

      return this.searchAllScopes(args.options);
    })()
      .then((customAction: CustomAction): void => {
        if (customAction["odata.null"] === true) {
          if (this.verbose) {
            cmd.log(`Custom action with id ${args.options.id} not found`);
          }
        }
        else {
          cmd.log({
            ClientSideComponentId: customAction.ClientSideComponentId,
            ClientSideComponentProperties: customAction.ClientSideComponentProperties,
            CommandUIExtension: customAction.CommandUIExtension,
            Description: customAction.Description,
            Group: customAction.Group,
            Id: customAction.Id,
            ImageUrl: customAction.ImageUrl,
            Location: customAction.Location,
            Name: customAction.Name,
            RegistrationId: customAction.RegistrationId,
            RegistrationType: customAction.RegistrationType,
            Rights: JSON.stringify(customAction.Rights),
            Scope: this.humanizeScope(customAction.Scope),
            ScriptBlock: customAction.ScriptBlock,
            ScriptSrc: customAction.ScriptSrc,
            Sequence: customAction.Sequence,
            Title: customAction.Title,
            Url: customAction.Url,
            VersionOfUserCustomAction: customAction.VersionOfUserCustomAction
          });
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  private getCustomAction(options: Options): Promise<CustomAction> {
    const requestOptions: any = {
      url: `${options.url}/_api/${options.scope}/UserCustomActions('${encodeURIComponent(options.id)}')`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      json: true
    };

    return request.get(requestOptions);
  }

  /**
   * Get request with `web` scope is send first. 
   * If custom action not found then 
   * another get request is send with `site` scope.
   */
  private searchAllScopes(options: Options): Promise<CustomAction> {
    return new Promise<CustomAction>((resolve: (customAction: CustomAction) => void, reject: (error: any) => void): void => {
      options.scope = "Web";

      this
        .getCustomAction(options)
        .then((webResult: CustomAction): void => {
          if (webResult["odata.null"] !== true) {
            return resolve(webResult);
          }

          options.scope = "Site";
          this
            .getCustomAction(options)
            .then((siteResult: CustomAction): void => {
              return resolve(siteResult);
            }, (err: any): void => {
              reject(err);
            });
        }, (err: any): void => {
          reject(err);
        });
    });
  }

  private humanizeScope(scope: number): string {
    switch (scope) {
      case 2:
        return "Site";
      case 3:
        return "Web";
    }

    return `${scope}`;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>',
        description: 'Id (Guid) of the custom action to retrieve'
      },
      {
        option: '-u, --url <url>',
        description: 'Url of the site (collection) to retrieve the custom action from'
      },
      {
        option: '-s, --scope [scope]',
        description: 'Scope of the custom action. Allowed values Site|Web|All. Default All',
        autocomplete: ['Site', 'Web', 'All']
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (Utils.isValidGuid(args.options.id) === false) {
        return `${args.options.id} is not valid. Custom action id (Guid) expected.`;
      }

      if (SpoCommand.isValidSharePointUrl(args.options.url) !== true) {
        return 'Missing required option url';
      }

      if (args.options.scope) {
        if (args.options.scope !== 'Site' &&
          args.options.scope !== 'Web' &&
          args.options.scope !== 'All') {
          return `${args.options.scope} is not a valid custom action scope. Allowed values are Site|Web|All`;
        }
      }

      return true;
    };
  }
}

module.exports = new SpoCustomActionGetCommand();