import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import { BasePermissions, PermissionKind } from '../../base-permissions';
import commands from '../../commands';
import { CustomAction } from './customaction';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  id: string;
  name?: string;
  title?: string;
  location?: string;
  group?: string;
  description?: string;
  sequence?: number;
  actionUrl?: string;
  imageUrl?: string;
  commandUIExtension?: string;
  registrationId?: string;
  registrationType?: string;
  rights?: string;
  scriptSrc?: string;
  scriptBlock?: string;
  scope?: string;
  clientSideComponentId?: string;
  clientSideComponentProperties?: string;
}

class SpoCustomActionSetCommand extends SpoCommand {
  public get name(): string {
    return commands.CUSTOMACTION_SET;
  }

  public get description(): string {
    return 'Updates a user custom action for site or site collection';
  }

  /**
   * Maps the base PermissionsKind enum to string array so it can 
   * more easily be used in validation or descriptions.
   */
  protected get permissionsKindMap(): string[] {
    const result: string[] = [];

    for (const kind in PermissionKind) {
      if (typeof PermissionKind[kind] === 'number') {
        result.push(kind);
      }
    }
    return result;
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.location = args.options.location;
    telemetryProps.scope = args.options.scope || 'Web';
    telemetryProps.group = (!(!args.options.group)).toString();
    telemetryProps.description = (!(!args.options.description)).toString();
    telemetryProps.sequence = (!(!args.options.sequence)).toString();
    telemetryProps.actionUrl = (!(!args.options.actionUrl)).toString();
    telemetryProps.imageUrl = (!(!args.options.imageUrl)).toString();
    telemetryProps.commandUIExtension = (!(!args.options.commandUIExtension)).toString();
    telemetryProps.registrationId = args.options.registrationId;
    telemetryProps.registrationType = args.options.registrationType;
    telemetryProps.rights = args.options.rights;
    telemetryProps.scriptSrc = (!(!args.options.scriptSrc)).toString();
    telemetryProps.scriptBlock = (!(!args.options.scriptBlock)).toString();
    telemetryProps.clientSideComponentId = (!(!args.options.clientSideComponentId)).toString();
    telemetryProps.clientSideComponentProperties = (!(!args.options.clientSideComponentProperties)).toString();
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    ((): Promise<CustomAction | undefined> => {
      if (!args.options.scope) {
        args.options.scope = 'All';
      }

      if (args.options.scope.toLowerCase() !== "all") {
        return this.updateCustomAction(args.options);
      }

      return this.searchAllScopes(args.options);
    })()
      .then((customAction: CustomAction | undefined): void => {
        if (this.verbose) {
          if (customAction && customAction["odata.null"] === true) {
            logger.logToStderr(`Custom action with id ${args.options.id} not found`);
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>'
      },
      {
        option: '-i, --id <id>'
      },
      {
        option: '-n, --name [name]'
      },
      {
        option: '-t, --title [title]'
      },
      {
        option: '-l, --location [location]'
      },
      {
        option: '-g, --group [group]'
      },
      {
        option: '-d, --description [description]'
      },
      {
        option: '--sequence [sequence]'
      },
      {
        option: '--actionUrl [actionUrl]'
      },
      {
        option: '--imageUrl [imageUrl]'
      },
      {
        option: '-e, --commandUIExtension [commandUIExtension]'
      },
      {
        option: '--registrationId [registrationId]'
      },
      {
        option: '--registrationType [registrationType]',
        autocomplete: ['None', 'List', 'ContentType', 'ProgId', 'FileType']
      },
      {
        option: '--rights [rights]',
        autocomplete: this.permissionsKindMap
      },
      {
        option: '-s, --scope [scope]',
        autocomplete: ['Site', 'Web', 'All']
      },
      {
        option: '--scriptBlock [scriptBlock]'
      },
      {
        option: '--scriptSrc [scriptSrc]'
      },
      {
        option: '-c, --clientSideComponentId [clientSideComponentId]'
      },
      {
        option: '-p, --clientSideComponentProperties [clientSideComponentProperties]'
      }
    ];
    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (validation.isValidGuid(args.options.id) === false) {
      return `${args.options.id} is not valid. Custom action id (Guid) expected`;
    }

    const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.url);
    if (isValidSharePointUrl !== true) {
      return isValidSharePointUrl;
    }

    if (!args.options.title && !args.options.name && !args.options.location &&
      !args.options.actionUrl && !args.options.clientSideComponentId && !args.options.clientSideComponentProperties &&
      !args.options.commandUIExtension && !args.options.group && !args.options.imageUrl &&
      !args.options.description && !args.options.registrationId && !args.options.registrationType &&
      !args.options.rights && !args.options.scriptBlock && !args.options.scriptSrc &&
      !args.options.sequence) {
      return 'Please specify option to be updated';
    }

    if (args.options.scriptSrc && args.options.scriptBlock) {
      return 'Either option scriptSrc or scriptBlock can be specified, but not both';
    }

    if (args.options.sequence && (args.options.sequence < 0 || args.options.sequence > 65536)) {
      return 'Invalid option sequence. Expected value in range from 0 to 65536';
    }

    if (args.options.clientSideComponentId && validation.isValidGuid(args.options.clientSideComponentId) === false) {
      return `ClientSideComponentId ${args.options.clientSideComponentId} is not a valid GUID`;
    }

    if (args.options.scope &&
      args.options.scope !== 'Site' &&
      args.options.scope !== 'Web' &&
      args.options.scope !== 'All'
    ) {
      return `${args.options.scope} is not a valid custom action scope. Allowed values are Site|Web|All`;
    }

    if (args.options.rights) {
      const rights = args.options.rights.split(',');

      for (const item of rights) {
        const kind: PermissionKind = PermissionKind[(item.trim() as keyof typeof PermissionKind)];

        if (!kind) {
          return `Rights option '${item}' is not recognized as valid PermissionKind choice. Please note it is case-sensitive`;
        }
      }
    }

    return true;
  }

  private updateCustomAction(options: Options): Promise<undefined> {
    const requestBody: any = this.mapRequestBody(options);

    const requestOptions: any = {
      url: `${options.url}/_api/${options.scope}/UserCustomActions('${encodeURIComponent(options.id)}')`,
      headers: {
        accept: 'application/json;odata=nometadata',
        'X-HTTP-Method': 'MERGE'
      },
      data: requestBody,
      responseType: 'json'
    };

    return request.post(requestOptions);
  }

  /**
   * Merge request with `web` scope is send first. 
   * If custom action not found then 
   * another merge request is send with `site` scope.
   */
  private searchAllScopes(options: Options): Promise<CustomAction | undefined> {
    return new Promise<CustomAction | undefined>((resolve: (customAction: CustomAction | undefined) => void, reject: (error: any) => void): void => {
      options.scope = "Web";

      this
        .updateCustomAction(options)
        .then((webResult: CustomAction | undefined): void => {
          if (webResult === undefined || webResult["odata.null"] !== true) {
            return resolve(webResult);
          }

          options.scope = "Site";
          this
            .updateCustomAction(options)
            .then((siteResult: CustomAction | undefined): void => {
              return resolve(siteResult);
            }, (err: any): void => {
              reject(err);
            });
        }, (err: any): void => {
          reject(err);
        });
    });
  }

  private mapRequestBody(options: Options): any {
    const requestBody: any = {};

    if (options.location) {
      requestBody.Location = options.location;
    }

    if (options.name) {
      requestBody.Name = options.name;
    }

    if (options.title) {
      requestBody.Title = options.title;
    }

    if (options.group) {
      requestBody.Group = options.group;
    }

    if (options.description) {
      requestBody.Description = options.description;
    }

    if (options.sequence) {
      requestBody.Sequence = options.sequence;
    }

    if (options.registrationType) {
      requestBody.RegistrationType = this.getRegistrationType(options.registrationType);
    }

    if (options.registrationId) {
      requestBody.RegistrationId = options.registrationId.toString();
    }

    if (options.actionUrl) {
      requestBody.Url = options.actionUrl;
    }

    if (options.imageUrl) {
      requestBody.ImageUrl = options.imageUrl;
    }

    if (options.clientSideComponentId) {
      requestBody.ClientSideComponentId = options.clientSideComponentId;
    }

    if (options.clientSideComponentProperties) {
      requestBody.ClientSideComponentProperties = options.clientSideComponentProperties;
    }

    if (options.scriptBlock) {
      requestBody.ScriptBlock = options.scriptBlock;
    }

    if (options.scriptSrc) {
      requestBody.ScriptSrc = options.scriptSrc;
    }

    if (options.commandUIExtension) {
      requestBody.CommandUIExtension = `${options.commandUIExtension}`;
    }

    if (options.rights) {
      const permissions: BasePermissions = new BasePermissions();
      const rights = options.rights.split(',');

      for (const item of rights) {
        const kind: PermissionKind = PermissionKind[(item.trim() as keyof typeof PermissionKind)];

        permissions.set(kind);
      }
      requestBody.Rights = {
        High: permissions.high.toString(),
        Low: permissions.low.toString()
      };
    }

    return requestBody;
  }

  private getRegistrationType(registrationType: string): number {
    switch (registrationType.toLowerCase()) {
      case 'list':
        return 1;
      case 'contenttype':
        return 2;
      case 'progid':
        return 3;
      case 'filetype':
        return 4;
    }
    return 0; // None
  }
}

module.exports = new SpoCustomActionSetCommand();