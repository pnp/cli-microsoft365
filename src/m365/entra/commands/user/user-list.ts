import { User } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { optionsUtils } from '../../../../utils/optionsUtils.js';
import { zod } from '../../../../utils/zod.js';

const allowedTypes = { Member: 'Member', Guest: 'Guest' } as const;

export const options = z.looseObject({
  ...globalOptionsZod.shape,
  type: zod.coercedEnum(allowedTypes).optional(),
  properties: z.string().optional().alias('p')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraUserListCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_LIST;
  }

  public get description(): string {
    return 'Lists users matching specified criteria';
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'mail', 'userPrincipalName'];
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let url = `${this.resource}/v1.0/users`;

      if (args.options.properties) {
        const selectProperties = args.options.properties;
        const allSelectProperties = selectProperties.split(',');
        const propertiesWithSlash = allSelectProperties.filter(item => item.includes('/'));

        const fieldExpand = propertiesWithSlash
          .map(p => `${p.split('/')[0]}($select=${p.split('/')[1]})`)
          .join(',');

        const expandParam = fieldExpand.length > 0 ? `&$expand=${fieldExpand}` : '';
        const selectParam = allSelectProperties.filter(item => !item.includes('/'));

        url += `?$select=${selectParam}${expandParam}`;
      }

      const filter = this.getFilter(args.options);
      if (filter) {
        url += `${args.options.properties ? '&' : '?'}${filter}`;
      }

      const users = await odata.getAllItems<User>(url);
      await logger.log(users);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getFilter(options: Options): string | null {
    const filters: string[] = [];

    const unknownOptions = optionsUtils.getUnknownOptions(options, zod.schemaToOptions(this.schema!));
    Object.keys(unknownOptions).forEach(key => {
      if (typeof (options as any)[key] === 'boolean') {
        throw `Specify value for the ${key} property`;
      }

      filters.push(`startsWith(${key}, '${formatting.encodeQueryParameter((options as any)[key].toString())}')`);
    });

    if (options.type) {
      filters.push(`userType eq '${options.type}'`);
    }

    if (filters.length > 0) {
      return `$filter=${filters.join(' and ')}`;
    }

    return null;
  }
}

export default new EntraUserListCommand();
