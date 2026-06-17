import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.uuid().optional().alias('i'),
  displayName: z.string().optional().alias('n'),
  objectId: z.uuid().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraEnterpriseAppGetCommand extends GraphCommand {
  public get name(): string {
    return commands.ENTERPRISEAPP_GET;
  }

  public get description(): string {
    return 'Gets information about an Enterprise Application';
  }

  public alias(): string[] | undefined {
    return [commands.SP_GET];
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

  private async getSpId(args: CommandArgs): Promise<string> {
    if (args.options.objectId) {
      return args.options.objectId;
    }

    let spMatchQuery: string = '';
    if (args.options.displayName) {
      spMatchQuery = `displayName eq '${formatting.encodeQueryParameter(args.options.displayName)}'`;
    }
    else if (args.options.id) {
      spMatchQuery = `appId eq '${formatting.encodeQueryParameter(args.options.id)}'`;
    }

    const idRequestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/servicePrincipals?$filter=${spMatchQuery}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: { id: string; }[] }>(idRequestOptions);

    const spItem: { id: string } | undefined = response.value[0];

    if (!spItem) {
      throw `The specified Entra app does not exist`;
    }

    if (response.value.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', response.value);
      const result = await cli.handleMultipleResultsFound<{ id: string }>(`Multiple Entra apps with name '${args.options.displayName}' found.`, resultAsKeyValuePair);
      return result.id;
    }

    return spItem.id;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving enterprise application information...`);
    }

    try {
      const id = await this.getSpId(args);

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/servicePrincipals/${id}`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const res = await request.get(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraEnterpriseAppGetCommand();