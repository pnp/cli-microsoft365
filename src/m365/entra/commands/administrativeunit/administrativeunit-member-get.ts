import { DirectoryObject } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { Logger } from '../../../../cli/Logger.js';
import { entraAdministrativeUnit } from '../../../../utils/entraAdministrativeUnit.js';
import request, { CliRequestOptions } from '../../../../request.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.uuid().alias('i'),
  administrativeUnitId: z.uuid().optional().alias('u'),
  administrativeUnitName: z.string().optional().alias('n'),
  properties: z.string().optional().alias('p')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

interface DirectoryObjectEx extends DirectoryObject {
  '@odata.context'?: string;
  '@odata.type'?: string;
  type: string;
}

class EntraAdministrativeUnitMemberGetCommand extends GraphCommand {
  public get name(): string {
    return commands.ADMINISTRATIVEUNIT_MEMBER_GET;
  }

  public get description(): string {
    return 'Retrieves info about a specific member of an administrative unit';
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
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let administrativeUnitId = args.options.administrativeUnitId;

    try {
      if (args.options.administrativeUnitName) {
        if (this.verbose) {
          await logger.logToStderr(`Retrieving Administrative Unit Id...`);
        }

        administrativeUnitId = (await entraAdministrativeUnit.getAdministrativeUnitByDisplayName(args.options.administrativeUnitName)).id!;
      }

      const url = this.getRequestUrl(administrativeUnitId!, args.options.id, args.options);

      const requestOptions: CliRequestOptions = {
        url: url,
        headers: {
          accept: 'application/json;odata.metadata=minimal'
        },
        responseType: 'json'
      };

      const result = await request.get<DirectoryObjectEx>(requestOptions);
      const odataType = result['@odata.type'];

      if (odataType) {
        result.type = odataType.replace('#microsoft.graph.', '');
      }

      delete result['@odata.type'];
      delete result['@odata.context'];

      await logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getRequestUrl(administrativeUnitId: string, memberId: string, options: Options): string {
    const queryParameters: string[] = [];

    if (options.properties) {
      const allProperties = options.properties.split(',');
      const selectProperties = allProperties.filter(prop => !prop.includes('/'));
      const expandProperties = allProperties.filter(prop => prop.includes('/'));

      if (selectProperties.length > 0) {
        queryParameters.push(`$select=${selectProperties}`);
      }

      if (expandProperties.length > 0) {
        const fieldExpands = expandProperties.map(p => {
          const properties = p.split('/');
          return `${properties[0]}($select=${properties[1]})`;
        });
        queryParameters.push(`$expand=${fieldExpands.join(',')}`);
      }
    }

    const queryString = queryParameters.length > 0
      ? `?${queryParameters.join('&')}`
      : '';

    return `${this.resource}/v1.0/directory/administrativeUnits/${administrativeUnitId}/members/${memberId}${queryString}`;
  }
}

export default new EntraAdministrativeUnitMemberGetCommand();