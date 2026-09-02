import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spe } from '../../../../utils/spe.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.uuid().optional().alias('i'),
  name: z.string().optional().alias('n'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpeContainerTypeRemoveCommand extends GraphDelegatedCommand {

  public get name(): string {
    return commands.CONTAINERTYPE_REMOVE;
  }

  public get description(): string {
    return 'Removes a specific container type';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.id, options.name].filter(o => o !== undefined).length === 1, {
        error: 'Use one of the following options: id, name.',
        params: {
          customCode: 'optionSet',
          options: ['id', 'name']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (!args.options.force) {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove container type ${args.options.id || args.options.name}?` });

      if (!result) {
        return;
      }
    }

    try {
      const containerTypeId = await this.getContainerTypeId(args.options, logger);

      if (this.verbose) {
        await logger.logToStderr(`Removing container type '${args.options.id || args.options.name}'...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/storage/fileStorage/containerTypes/${containerTypeId}`,
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

  private async getContainerTypeId(options: Options, logger: Logger): Promise<string> {
    if (options.id) {
      return options.id;
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving container type id for container type '${options.name}'...`);
    }

    return spe.getContainerTypeIdByName(options.name!);
  }
}

export default new SpeContainerTypeRemoveCommand();