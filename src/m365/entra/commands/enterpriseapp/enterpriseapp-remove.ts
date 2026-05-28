import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { odata } from '../../../../utils/odata.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.uuid().optional().alias('i'),
  displayName: z.string().optional().alias('n'),
  objectId: z.uuid().optional(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraEnterpriseAppRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.ENTERPRISEAPP_REMOVE;
  }

  public get description(): string {
    return 'Deletes an enterprise application (or service principal)';
  }

  public alias(): string[] | undefined {
    return [commands.SP_REMOVE];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.id, options.displayName, options.objectId].filter(o => o !== undefined).length === 1, {
        error: 'Specify either id, displayName, or objectId',
        params: {
          customCode: 'optionSet',
          options: ['id', 'displayName', 'objectId']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeEnterpriseApplication = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Removing enterprise application ${args.options.id || args.options.displayName || args.options.objectId}...`);
      }

      try {
        let url = `${this.resource}/v1.0`;

        if (args.options.id) {
          url += `/servicePrincipals(appId='${args.options.id}')`;
        }
        else {
          const id = await this.getSpId(args.options);
          url += `/servicePrincipals/${id}`;
        }

        const requestOptions: CliRequestOptions = {
          url: url,
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
    };

    if (args.options.force) {
      await removeEnterpriseApplication();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove enterprise application '${args.options.id || args.options.displayName || args.options.objectId}'?` });

      if (result) {
        await removeEnterpriseApplication();
      }
    }
  }

  private async getSpId(options: Options): Promise<string> {
    if (options.objectId) {
      return options.objectId;
    }

    const spItemsResponse = await odata.getAllItems<{ id: string }>(`${this.resource}/v1.0/servicePrincipals?$filter=displayName eq '${formatting.encodeQueryParameter(options.displayName!)}'&$select=id`);

    if (spItemsResponse.length === 0) {
      throw `The specified enterprise application does not exist.`;
    }

    if (spItemsResponse.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', spItemsResponse);
      const result = await cli.handleMultipleResultsFound<{ id: string }>(`Multiple enterprise applications with name '${options.displayName}' found.`, resultAsKeyValuePair);
      return result.id;
    }

    const spItem = spItemsResponse[0];

    return spItem.id;
  }
}

export default new EntraEnterpriseAppRemoveCommand();