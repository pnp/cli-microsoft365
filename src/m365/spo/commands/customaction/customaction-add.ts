import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import { BasePermissions, PermissionKind } from '../../base-permissions.js';
import commands from '../../commands.js';
import { CustomAction } from './customaction.js';
import { z } from 'zod';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string().refine(val => validation.isValidSharePointUrl(val) === true, {
    error: e => `${e.input} is not a valid SharePoint Online site URL.`
  }).alias('u'),
  name: z.string().alias('n'),
  title: z.string().alias('t'),
  location: z.string().alias('l'),
  group: z.string().optional().alias('g'),
  description: z.string().optional().alias('d'),
  sequence: z.coerce.number().refine(val => val >= 0 && val <= 65536, {
    message: 'Invalid option sequence. Expected value in range from 0 to 65536'
  }).optional(),
  actionUrl: z.string().optional(),
  imageUrl: z.string().optional(),
  commandUIExtension: z.string().optional().alias('e'),
  registrationId: z.string().optional(),
  registrationType: z.preprocess(value => {
    const values = ['None', 'List', 'ContentType', 'ProgId', 'FileType'];
    return typeof value === 'string'
      ? values.find(type => type.toLowerCase() === value.toLowerCase()) ?? value
      : value;
  }, z.enum(['None', 'List', 'ContentType', 'ProgId', 'FileType'])).optional(),
  rights: z.string().refine(val => {
    const items = val.split(',').map(s => s.trim());
    return items.every(item => PermissionKind[(item as keyof typeof PermissionKind)] !== undefined);
  }, {
    error: e => {
      const items = (e.input as string).split(',').map((s: string) => s.trim());
      const invalid = items.find((item: string) => !PermissionKind[(item as keyof typeof PermissionKind)]);
      return `Rights option '${invalid}' is not recognized as valid PermissionKind choice. Please note it is case sensitive`;
    }
  }).optional(),
  scriptSrc: z.string().optional(),
  scriptBlock: z.string().optional(),
  scope: z.enum(['Site', 'Web']).optional().alias('s'),
  clientSideComponentId: z.string().refine(val => validation.isValidGuid(val), {
    error: e => `ClientSideComponentId ${e.input} is not a valid GUID`
  }).optional().alias('c'),
  clientSideComponentProperties: z.string().optional().alias('p')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoCustomActionAddCommand extends SpoCommand {
  public get name(): string {
    return commands.CUSTOMACTION_ADD;
  }

  public get description(): string {
    return 'Adds a user custom action for site or site collection';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => !opts.registrationId || opts.registrationType, {
        message: 'Option registrationId is specified, but registrationType is missing',
        path: ['registrationType'],
        params: { customCode: 'required' }
      })
      .refine(opts => !opts.registrationType || opts.registrationId, {
        message: 'Option registrationType is specified, but registrationId is missing',
        path: ['registrationId'],
        params: { customCode: 'required' }
      })
      .refine(opts => {
        if (!opts.location) {
          return true;
        }
        const location = opts.location.toLowerCase();
        const locationsRequireGroup = [
          'microsoft.sharepoint.standardmenu', 'microsoft.sharepoint.contenttypesettings',
          'microsoft.sharepoint.contenttypetemplatesettings', 'microsoft.sharepoint.create',
          'microsoft.sharepoint.groupspage', 'microsoft.sharepoint.listedit',
          'microsoft.sharepoint.listedit.documentlibrary', 'microsoft.sharepoint.peoplepage',
          'microsoft.sharepoint.sitesettings'
        ];
        return !locationsRequireGroup.includes(location) || !!opts.group;
      }, {
        message: 'The location specified requires the group option to be specified as well',
        path: ['group'],
        params: { customCode: 'required' }
      })
      .refine(opts => {
        if (!opts.location) {
          return true;
        }
        const location = opts.location.toLowerCase();
        if (location === 'scriptlink') {
          return !!opts.scriptSrc || !!opts.scriptBlock;
        }
        return true;
      }, {
        message: 'Option scriptSrc or scriptBlock is required when the location is set to ScriptLink',
        params: {
          customCode: 'required',
          options: ['scriptSrc', 'scriptBlock']
        }
      })
      .refine(opts => {
        if (!opts.location) {
          return true;
        }
        const location = opts.location.toLowerCase();
        if ((opts.scriptSrc || opts.scriptBlock) && location !== 'scriptlink') {
          return false;
        }
        return true;
      }, {
        message: 'Option scriptSrc or scriptBlock is specified, but the location option is different than ScriptLink. Please use --actionUrl, if the location should be different than ScriptLink'
      })
      .refine(opts => !(opts.scriptSrc && opts.scriptBlock), {
        message: 'Either option scriptSrc or scriptBlock can be specified, but not both'
      })
      .refine(opts => !opts.clientSideComponentProperties || opts.clientSideComponentId, {
        message: 'Option clientSideComponentProperties is specified, but the clientSideComponentId option is missing',
        path: ['clientSideComponentId'],
        params: { customCode: 'required' }
      });
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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (!args.options.scope) {
        args.options.scope = 'Web';
      }

      const requestBody: any = this.mapRequestBody(args.options);

      const requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/${args.options.scope}/UserCustomActions`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        data: requestBody,
        responseType: 'json'
      };

      const customAction = await request.post<CustomAction>(requestOptions);
      if (this.verbose) {
        await logger.logToStderr({
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
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
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

export default new SpoCustomActionAddCommand();