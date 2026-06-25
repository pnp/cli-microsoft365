import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { odata } from '../../../../utils/odata.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  environmentName: z.string().alias('e'),
  entitySetName: z.string().optional(),
  tableName: z.string().optional(),
  asAdmin: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PpDataverseTableRowListCommand extends PowerPlatformCommand {

  public get name(): string {
    return commands.DATAVERSE_TABLE_ROW_LIST;
  }

  public get description(): string {
    return 'Lists table rows for the given Dataverse table';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.entitySetName, opts.tableName].filter(x => x !== undefined).length === 1, {
        message: `Specify either 'entitySetName' or 'tableName', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['entitySetName', 'tableName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Retrieving list of table rows');
    }

    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);

      const entitySetName = await this.getEntitySetName(dynamicsApiUrl, args);
      if (this.verbose) {
        await logger.logToStderr('Entity set name is: ' + entitySetName);
      }

      const response = await odata.getAllItems<any>(`${dynamicsApiUrl}/api/data/v9.0/${entitySetName}`);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  protected async getEntitySetName(dynamicsApiUrl: string, args: CommandArgs): Promise<string> {
    if (args.options.entitySetName) {
      return args.options.entitySetName;
    }

    const requestOptions: CliRequestOptions = {
      url: `${dynamicsApiUrl}/api/data/v9.0/EntityDefinitions(LogicalName='${args.options.tableName}')?$select=EntitySetName`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ EntitySetName: string }>(requestOptions);

    return response.EntitySetName;
  }
}

export default new PpDataverseTableRowListCommand();