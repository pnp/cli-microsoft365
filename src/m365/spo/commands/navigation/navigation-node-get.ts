import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  id: z.int().positive()
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoNavigationNodeGetCommand extends SpoCommand {
  public get name(): string {
    return commands.NAVIGATION_NODE_GET;
  }

  public get description(): string {
    return 'Retrieve information about a specific navigation node';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving information about navigation node with id ${args.options.id}`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${args.options.webUrl}/_api/web/navigation/GetNodeById(${args.options.id})?$expand=Children,Children/Children,Children/Children/Children`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const listInstance = await request.get<any>(requestOptions);
      if (listInstance['odata.null']) {
        throw `No navigation node found with id ${args.options.id}.`;
      }

      await logger.log(listInstance);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoNavigationNodeGetCommand();