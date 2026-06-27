import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().alias('i')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpeContainerActivateCommand extends GraphCommand {
  public get name(): string {
    return commands.CONTAINER_ACTIVATE;
  }

  public get description(): string {
    return 'Activates a container';
  }

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Activating a container with id '${args.options.id}'...`);
    }

    try {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/storage/fileStorage/containers/${formatting.encodeQueryParameter(args.options.id)}/activate`,
        headers: {
          'content-type': 'application/json;odata=nometadata',
          'accept': 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpeContainerActivateCommand();