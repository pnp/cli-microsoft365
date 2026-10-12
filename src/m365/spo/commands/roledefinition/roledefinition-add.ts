import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import { BasePermissions, PermissionKind } from '../../base-permissions.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  name: z.string().alias('n'),
  description: z.string().optional().alias('d'),
  rights: z.string().optional()
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoRoleDefinitionAddCommand extends SpoCommand {
  public get name(): string {
    return commands.ROLEDEFINITION_ADD;
  }

  public get description(): string {
    return 'Adds a new roledefinition to web';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  protected get permissionsKindMap(): string[] {
    const result: string[] = [];

    for (const kind in PermissionKind) {
      if (typeof PermissionKind[kind] === 'number') {
        result.push(kind);
      }
    }
    return result;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine((opts: { rights?: string }) => {
        if (!opts.rights) {
          return true;
        }

        const rights = opts.rights.split(',');
        for (const item of rights) {
          const kind: PermissionKind = PermissionKind[(item.trim() as keyof typeof PermissionKind)];
          if (!kind) {
            return false;
          }
        }
        return true;
      }, {
        error: (e: any) => {
          const rights = (e.input as { rights?: string }).rights!.split(',');
          const invalid = rights.filter((item: string) => {
            const kind: PermissionKind = PermissionKind[(item.trim() as keyof typeof PermissionKind)];
            return !kind;
          });
          return {
            message: `Rights option '${invalid[0]}' is not recognized as valid PermissionKind choice. Please note it is case-sensitive. Allowed values are ${this.permissionsKindMap.join('|')}.`
          };
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Adding role definition to ${args.options.webUrl}...`);
    }

    const description = args.options.description || '';

    const permissions: BasePermissions = new BasePermissions();
    if (args.options.rights) {
      const rights = args.options.rights.split(',');

      for (const item of rights) {
        const kind: PermissionKind = PermissionKind[(item.trim() as keyof typeof PermissionKind)];

        permissions.set(kind);
      }
    }

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web/roledefinitions`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json',
      data: {
        'BasePermissions': {
          'High': permissions.high.toString(),
          'Low': permissions.low.toString()
        },
        'Description': `${description}`,
        'Name': `${args.options.name}`
      }
    };

    try {
      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoRoleDefinitionAddCommand();