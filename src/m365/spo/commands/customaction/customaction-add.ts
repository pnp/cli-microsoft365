import { Logger } from '../../../../cli';
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
  name: string;
  title: string;
  location: string;
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

class SpoCustomActionAddCommand extends SpoCommand {
  public get name(): string {
    return commands.CUSTOMACTION_ADD;
  }

  public get description(): string {
    return 'Adds a user custom action for site or site collection';
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
        location: args.options.location,
        scope: args.options.scope || 'Web',
        group: (!(!args.options.group)).toString(),
        description: (!(!args.options.description)).toString(),
        sequence: (!(!args.options.sequence)).toString(),
        actionUrl: (!(!args.options.actionUrl)).toString(),
        imageUrl: (!(!args.options.imageUrl)).toString(),
        commandUIExtension: (!(!args.options.commandUIExtension)).toString(),
        registrationId: args.options.registrationId,
        registrationType: args.options.registrationType,
        rights: args.options.rights,
        scriptSrc: (!(!args.options.scriptSrc)).toString(),
        scriptBlock: (!(!args.options.scriptBlock)).toString(),
        clientSideComponentId: (!(!args.options.clientSideComponentId)).toString(),
        clientSideComponentProperties: (!(!args.options.clientSideComponentProperties)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --url <url>'
      },
      {
        option: '-n, --name <name>'
      },
      {
        option: '-t, --title <title>'
      },
      {
        option: '-l, --location <location>'
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
        autocomplete: ['Site', 'Web']
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
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.url);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }
    
        if (args.options.registrationId && !args.options.registrationType) {
          return 'Option registrationId is specified, but registrationType is missing';
        }
    
        if (args.options.registrationType && !args.options.registrationId) {
          return 'Option registrationType is specified, but registrationId is missing';
        }
    
        const location: string = args.options.location.toLowerCase();
        const locationsRequireGroup: string[] = [
          'microsoft.sharepoint.standardmenu', 'microsoft.sharepoint.contenttypesettings',
          'microsoft.sharepoint.contenttypetemplatesettings', 'microsoft.sharepoint.create',
          'microsoft.sharepoint.groupspage', 'microsoft.sharepoint.listedit',
          'microsoft.sharepoint.listedit.documentlibrary', 'microsoft.sharepoint.peoplepage',
          'microsoft.sharepoint.sitesettings'
        ];
    
        if (locationsRequireGroup.indexOf(location) > -1 && !args.options.group) {
          return `The location specified requires the group option to be specified as well`;
        }
    
        if (location === 'scriptlink' &&
          !args.options.scriptSrc &&
          !args.options.scriptBlock
        ) {
          return 'Option scriptSrc or scriptBlock is required when the location is set to ScriptLink';
        }
    
        if ((args.options.scriptSrc || args.options.scriptBlock) && location !== 'scriptlink') {
          return 'Option scriptSrc or scriptBlock is specified, but the location option is different than ScriptLink. Please use --actionUrl, if the location should be different than ScriptLink';
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
    
        if (args.options.clientSideComponentProperties && !args.options.clientSideComponentId) {
          return `Option clientSideComponentProperties is specified, but the clientSideComponentId option is missing`;
        }
    
        if (args.options.scope &&
          args.options.scope !== 'Site' &&
          args.options.scope !== 'Web'
        ) {
          return `${args.options.scope} is not a valid custom action scope. Allowed values are Site|Web`;
        }
    
        if (args.options.rights) {
          const rights = args.options.rights.split(',');
    
          for (const item of rights) {
            const kind: PermissionKind = PermissionKind[(item.trim() as keyof typeof PermissionKind)];
    
            if (!kind) {
              return `Rights option '${item}' is not recognized as valid PermissionKind choice. Please note it is case sensitive`;
            }
          }
        }
    
        return true;
      }
    );
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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (!args.options.scope) {
      args.options.scope = 'Web';
    }

    const requestBody: any = this.mapRequestBody(args.options);

    const requestOptions: any = {
      url: `${args.options.url}/_api/${args.options.scope}/UserCustomActions`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      data: requestBody,
      responseType: 'json'
    };

    request
      .post<CustomAction>(requestOptions)
      .then((customAction: CustomAction): void => {
        if (this.verbose) {
          logger.logToStderr({
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
            Scope: args.options.scope, // because it is more human readable
            ScriptBlock: customAction.ScriptBlock,
            ScriptSrc: customAction.ScriptSrc,
            Sequence: customAction.Sequence,
            Title: customAction.Title,
            Url: customAction.Url,
            VersionOfUserCustomAction: customAction.VersionOfUserCustomAction
          });
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
  }

  private mapRequestBody(options: Options): any {
    const requestBody: any = {
      Title: options.title,
      Name: options.name,
      Location: options.location
    };

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

module.exports = new SpoCustomActionAddCommand();