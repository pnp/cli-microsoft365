import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  status: z.enum(['Available', 'InDevelopment', 'Deprecated']).optional().alias('s'),
  owner: z.string().optional(),
  pageSize: z.string().optional().alias('p'),
  pageNumber: z.string().optional().alias('n')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class GraphSchemaExtensionListCommand extends GraphCommand {
  public get name(): string {
    return commands.SCHEMAEXTENSION_LIST;
  }

  public get description(): string {
    return 'Get a list of schemaExtension objects created in the current tenant, that can be InDevelopment, Available, or Deprecated.';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => !options.owner || validation.isValidGuid(options.owner), {
        error: e => `${(e.input as Options).owner} is not a valid GUID`,
        path: ['owner']
      })
      .refine(options => !options.pageNumber || parseInt(options.pageNumber) >= 1, {
        error: 'pageNumber must be a positive number',
        path: ['pageNumber']
      })
      .refine(options => !options.pageSize || parseInt(options.pageSize) >= 1, {
        error: 'pageSize must be a positive number',
        path: ['pageSize']
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const filter: string = this.getFilter(args.options);
    let url = `${this.resource}/v1.0/schemaExtensions?$select=*${(filter.length > 0 ? '&' + filter : '')}`;

    if (args.options.pageNumber && Number(args.options.pageNumber) > 0) {
      const rowLimit: string = `&$top=${Number(args.options.pageSize ? args.options.pageSize : 10) * Number(args.options.pageNumber + 1)}`;
      url += rowLimit;
    }
    const requestOptions: CliRequestOptions = {
      url: url,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json'
      },
      responseType: 'json'
    };

    try {
      const res = await request.get<any>(requestOptions);
      if (res.value && res.value.length > 0) {
        const size = args.options.pageSize ? parseInt(args.options.pageSize) : parseInt(res.value.length);
        const result = res.value.slice(-size);
        if (args.options.output !== 'json' && result.length > 1) {
          await logger.log(result.map((x: any) => ({
            id: x.id,
            description: x.description,
            targetTypes: x.targetTypes,
            status: x.status,
            owner: x.owner,
            properties: JSON.stringify(x.properties)
          })));
        }
        else {
          await logger.log(result);
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }


  private getFilter(options: any): string {
    const filters: any = {};
    const filterOptions: string[] = [
      'status',
      'owner'
    ];

    Object.keys(options).forEach(key => {
      if (filterOptions.indexOf(key) !== -1) {
        filters[key] = options[key].replace(/'/g, `''`);
      }
    });
    let filter: string = Object.keys(filters).map(key => `${key} eq '${filters[key]}'`).join(' and ');
    if (filter.length > 0) {
      filter = '$filter=' + filter;
    }

    return filter;
  }
}
export default new GraphSchemaExtensionListCommand();