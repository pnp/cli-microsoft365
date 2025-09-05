import { Logger } from '../../../../cli/Logger.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { zod } from '../../../../utils/zod.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';
import { formatting } from '../../../../utils/formatting.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { odata } from '../../../../utils/odata.js';
import { cli } from '../../../../cli/cli.js';
import { SpeContainerType } from '../../../../utils/spe.js';

const options = globalOptionsZod
  .extend({
    id: zod.alias('i', z.string()
      .refine(id => validation.isValidGuid(id), id => ({
        message: `'${id}' is not a valid GUID.`
      }))
      .optional()
    ),
    name: zod.alias('n', z.string().optional())
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpeContainerTypeGetCommand extends GraphDelegatedCommand {
  public get name(): string {
    return commands.CONTAINERTYPE_GET;
  }

  public get description(): string {
    return 'Gets a container type';
  }

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => [options.id, options.name].filter(o => o !== undefined).length === 1, {
        message: 'Use one of the following options: id or name.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let result: SpeContainerType;
      if (args.options.name) {
        const containerTypes = await odata.getAllItems<SpeContainerType>(`${this.resource}/beta/storage/fileStorage/containerTypes?$filter=name eq '${formatting.encodeQueryParameter(args.options.name!)}'`);

        if (containerTypes.length === 0) {
          throw `The specified container type '${args.options.name}' does not exist.`;
        }

        if (containerTypes.length > 1) {
          const containerKeyValuePair = formatting.convertArrayToHashTable('id', containerTypes);
          result = await cli.handleMultipleResultsFound<SpeContainerType>(`Multiple container types with name '${args.options.name}' found.`, containerKeyValuePair);
        }
        else {
          result = containerTypes[0];
        }
      }
      else {
        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/beta/storage/fileStorage/containerTypes/${args.options.id}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        result = await request.get<SpeContainerType>(requestOptions);
      }

      await logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpeContainerTypeGetCommand();