import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import VivaEngageCommand from '../../../base/VivaEngageCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  withSuspended: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class VivaEngageNetworkListCommand extends VivaEngageCommand {
  public get name(): string {
    return commands.ENGAGE_NETWORK_LIST;
  }

  public get description(): string {
    return 'Returns a list of networks to which the current user has access';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'name', 'email', 'community', 'permalink', 'web_url'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1/networks/current.json`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json;odata=nometadata'
      },
      responseType: 'json',
      data: {
        includeSuspended: (args.options.withSuspended !== undefined && args.options.withSuspended !== false)
      }
    };

    try {
      const res: any = await request.get(requestOptions);

      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new VivaEngageNetworkListCommand();