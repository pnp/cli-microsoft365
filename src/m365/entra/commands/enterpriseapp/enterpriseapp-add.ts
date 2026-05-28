import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';

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

class EntraEnterpriseAppAddCommand extends GraphCommand {
  public get name(): string {
    return commands.ENTERPRISEAPP_ADD;
  }

  public get description(): string {
    return 'Creates an enterprise application (or service principal) for a registered Entra app';
  }

  public alias(): string[] | undefined {
    return [commands.SP_ADD];
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

  private async getAppId(args: CommandArgs): Promise<string> {
    if (args.options.id) {
      return args.options.id;
    }

    let spMatchQuery: string = '';
    if (args.options.displayName) {
      spMatchQuery = `displayName eq '${formatting.encodeQueryParameter(args.options.displayName)}'`;
    }
    else if (args.options.objectId) {
      spMatchQuery = `id eq '${formatting.encodeQueryParameter(args.options.objectId)}'`;
    }

    const appIdRequestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/applications?$filter=${spMatchQuery}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: { appId: string; }[] }>(appIdRequestOptions);

    const spItem: { appId: string } | undefined = response.value[0];

    if (!spItem) {
      throw `The specified Entra app doesn't exist`;
    }

    if (response.value.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('appId', response.value);
      const result = await cli.handleMultipleResultsFound<{ appId: string }>(`Multiple Entra apps with name '${args.options.displayName}' found.`, resultAsKeyValuePair);
      return result.appId;
    }

    return spItem.appId;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const appId = await this.getAppId(args);

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/servicePrincipals`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json;odata=nometadata'
        },
        data: {
          appId: appId
        },
        responseType: 'json'
      };

      const res = await request.post(requestOptions);

      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraEnterpriseAppAddCommand();