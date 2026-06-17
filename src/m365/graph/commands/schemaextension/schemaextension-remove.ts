import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().alias('i'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class GraphSchemaExtensionRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.SCHEMAEXTENSION_REMOVE;
  }

  public get description(): string {
    return 'Removes specified Microsoft Graph schema extension';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeSchemaExtension = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Removes specified Microsoft Graph schema extension with id '${args.options.id}'...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/schemaExtensions/${args.options.id}`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json'
        },
        responseType: 'json'
      };

      try {
        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeSchemaExtension();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the schema extension with ID ${args.options.id}?` });

      if (result) {
        await removeSchemaExtension();
      }
    }
  }
}
export default new GraphSchemaExtensionRemoveCommand();