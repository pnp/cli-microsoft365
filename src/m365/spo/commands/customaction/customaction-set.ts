import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
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
  id: z.string().refine(val => validation.isValidGuid(val), {
    error: e => `${e.input} is not valid. Custom action id (Guid) expected`
  }).alias('i'),
  name: z.string().optional().alias('n'),
  title: z.string().optional().alias('t'),
  location: z.string().optional().alias('l'),
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
      return `Rights option '${invalid}' is not recognized as valid PermissionKind choice. Please note it is case-sensitive`;
    }
  }).optional(),
  scriptSrc: z.string().optional(),
  scriptBlock: z.string().optional(),
  scope: z.enum(['Site', 'Web', 'All']).optional().alias('s'),
  clientSideComponentId: z.string().refine(val => validation.isValidGuid(val), {
    error: e => `ClientSideComponentId ${e.input} is not a valid GUID`
  }).optional().alias('c'),
  clientSideComponentProperties: z.string().optional().alias('p')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoCustomActionSetCommand extends SpoCommand {
  public get name(): string {
    return commands.CUSTOMACTION_SET;
  }

  public get description(): string {
    return 'Updates a user custom action for site or site collection';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => opts.title || opts.name || opts.location ||
        opts.actionUrl || opts.clientSideComponentId || opts.clientSideComponentProperties ||
        opts.commandUIExtension || opts.group || opts.imageUrl ||
        opts.description || opts.registrationId || opts.registrationType ||
        opts.rights || opts.scriptBlock || opts.scriptSrc ||
        opts.sequence, {
        message: 'Please specify option to be updated'
      })
      .refine(opts => !(opts.scriptSrc && opts.scriptBlock), {
        message: 'Either option scriptSrc or scriptBlock can be specified, but not both'
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
      let customAction: CustomAction | undefined;
      if (!args.options.scope) {
        args.options.scope = 'All';
      }
      if (args.options.scope.toLowerCase() !== "all") {
        customAction = await this.updateCustomAction(args.options);
      }
      else {
        customAction = await this.searchAllScopes(args.options);
      }

      if (this.verbose) {
        if (customAction && customAction["odata.null"] === true) {
          await logger.logToStderr(`Custom action with id ${args.options.id} not found`);
        }
      }
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }

  private async updateCustomAction(options: Options): Promise<CustomAction | undefined> {
    const requestBody: any = this.mapRequestBody(options);

    const requestOptions: CliRequestOptions = {
      url: `${options.webUrl}/_api/${options.scope}/UserCustomActions('${formatting.encodeQueryParameter(options.id)}')`,
      headers: {
        accept: 'application/json;odata=nometadata',
        'X-HTTP-Method': 'MERGE'
      },
      data: requestBody,
      responseType: 'json'
    };

    return await request.post(requestOptions);
  }

  /**
   * Merge request with `web` scope is send first. 
   * If custom action not found then 
   * another merge request is send with `site` scope.
   */
  private async searchAllScopes(options: Options): Promise<CustomAction | undefined> {
    options.scope = "Web";

    const webResult = await this.updateCustomAction(options);

    if (webResult === undefined || webResult["odata.null"] !== true) {
      return webResult;
    }

    options.scope = "Site";
    const siteResult = await this.updateCustomAction(options);
    return siteResult;
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

export default new SpoCustomActionSetCommand();