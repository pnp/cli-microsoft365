import { DirectoryObject } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { entraAdministrativeUnit } from '../../../../utils/entraAdministrativeUnit.js';
import { CliRequestOptions } from '../../../../request.js';
import { zod } from '../../../../utils/zod.js';

enum MemberType {
  User = 'user',
  Group = 'group',
  Device = 'device'
}

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  administrativeUnitId: z.uuid().optional().alias('i'),
  administrativeUnitName: z.string().optional().alias('n'),
  type: zod.coercedEnum(MemberType).optional().alias('t'),
  properties: z.string().optional().alias('p'),
  filter: z.string().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

interface DirectoryObjectEx extends DirectoryObject {
  '@odata.type'?: string;
  type: string;
}

class EntraAdministrativeUnitMemberListCommand extends GraphCommand {
  public get name(): string {
    return commands.ADMINISTRATIVEUNIT_MEMBER_LIST;
  }

  public get description(): string {
    return 'Retrieves members (users, groups, or devices) of an administrative unit.';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.administrativeUnitId, options.administrativeUnitName].filter(Boolean).length === 1, {
        error: 'Specify either administrativeUnitId or administrativeUnitName',
        params: {
          customCode: 'optionSet',
          options: ['administrativeUnitId', 'administrativeUnitName']
        }
      })
      .refine(options => !options.filter || options.type, {
        error: 'Filter can be specified only if type is set'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let administrativeUnitId = args.options.administrativeUnitId;

    try {
      if (args.options.administrativeUnitName) {
        administrativeUnitId = (await entraAdministrativeUnit.getAdministrativeUnitByDisplayName(args.options.administrativeUnitName)).id;
      }

      let results;
      const endpoint = this.getRequestUrl(administrativeUnitId!, args.options);

      if (args.options.type) {
        if (args.options.filter) {
          // While using the filter, we need to specify the ConsistencyLevel header.
          // Can be refactored when the header is no longer necessary.
          const requestOptions: CliRequestOptions = {
            url: endpoint,
            headers: {
              accept: 'application/json;odata.metadata=none',
              ConsistencyLevel: 'eventual'
            },
            responseType: 'json'
          };
          results = await odata.getAllItems<DirectoryObject>(requestOptions);
        }
        else {
          results = await odata.getAllItems<DirectoryObject>(endpoint);
        }
      }
      else {
        results = await odata.getAllItems<DirectoryObjectEx>(endpoint, 'minimal');

        results.forEach(c => {
          const odataType = c['@odata.type'];

          if (odataType) {
            c.type = odataType.replace('#microsoft.graph.', '');
          }

          delete c['@odata.type'];
        });
      }

      await logger.log(results);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getRequestUrl(administrativeUnitId: string, options: Options): string {
    const queryParameters: string[] = [];

    if (options.properties) {
      const allProperties = options.properties.split(',');
      const selectProperties = allProperties.filter(prop => !prop.includes('/'));
      const expandProperties = allProperties.filter(prop => prop.includes('/'));

      if (selectProperties.length > 0) {
        queryParameters.push(`$select=${selectProperties}`);
      }

      if (expandProperties.length > 0) {
        const fieldExpands = expandProperties.map(p => `${p.split('/')[0]}($select=${p.split('/')[1]})`);
        queryParameters.push(`$expand=${fieldExpands.join(',')}`);
      }
    }

    if (options.filter) {
      queryParameters.push(`$filter=${options.filter}`);
      queryParameters.push('$count=true');
    }

    const queryString = queryParameters.length > 0
      ? `?${queryParameters.join('&')}`
      : '';

    return options.type
      ? `${this.resource}/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/microsoft.graph.${options.type}${queryString}`
      : `${this.resource}/v1.0/directory/administrativeUnits/${administrativeUnitId}/members${queryString}`;
  }
}

export default new EntraAdministrativeUnitMemberListCommand();