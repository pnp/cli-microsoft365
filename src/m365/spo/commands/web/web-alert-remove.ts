import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string().alias('u')
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint URL.`
    }),
  id: z.uuid(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoWebAlertRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.WEB_ALERT_REMOVE;
  }

  public get description(): string {
    return 'Removes an alert from a SharePoint list';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (!args.options.force) {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the alert with id '${args.options.id}' from site '${args.options.webUrl}'?` });

      if (!result) {
        return;
      }
    }

    try {
      if (this.verbose) {
        await logger.logToStderr(`Removing alert with ID '${args.options.id}' from site '${args.options.webUrl}'...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/web/Alerts/DeleteAlert('${formatting.encodeQueryParameter(args.options.id)}')`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      await request.delete(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoWebAlertRemoveCommand();