import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { validation } from '../../../../utils/validation.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  environmentName: z.string().alias('e'),
  id: z.string().refine(val => validation.isValidGuid(val), {
    error: 'The value must be a valid GUID.'
  }).alias('i'),
  entitySetName: z.string().optional(),
  tableName: z.string().optional(),
  asAdmin: z.boolean().optional(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PpDataverseTableRowRemoveCommand extends PowerPlatformCommand {

  public get name(): string {
    return commands.DATAVERSE_TABLE_ROW_REMOVE;
  }

  public get description(): string {
    return 'Removes a specific row from a dataverse table in the specified Power Platform environment.';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.entitySetName, opts.tableName].filter(x => x !== undefined).length === 1, {
        error: `Specify either 'entitySetName' or 'tableName', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['entitySetName', 'tableName']
        }
      });
  }

  public async commandAction(logger: Logger, args: any): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Removing row '${args.options.id}' from table '${args.options.tableName || args.options.entitySetName}'...`);
    }

    if (args.options.force) {
      await this.deleteTableRow(logger, args);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove row '${args.options.id}' from table '${args.options.tableName || args.options.entitySetName}'?` });

      if (result) {
        await this.deleteTableRow(logger, args);
      }
    }
  }

  private async deleteTableRow(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);

      const entitySetName = await this.getEntitySetName(dynamicsApiUrl, args);
      if (this.verbose) {
        await logger.logToStderr('Entity set name is: ' + entitySetName);
      }

      const requestOptions: CliRequestOptions = {
        url: `${dynamicsApiUrl}/api/data/v9.1/${entitySetName}(${args.options.id})`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      await request.delete(requestOptions);
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

export default new PpDataverseTableRowRemoveCommand();