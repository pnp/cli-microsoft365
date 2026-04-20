import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string().alias('u')
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint URL.`
    }),
  id: z.uuid()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoWebAlertGetCommand extends SpoCommand {
  public get name(): string {
    return commands.WEB_ALERT_GET;
  }

  public get description(): string {
    return 'Retrieves details of a specific alert from a SharePoint site list';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving alert with id '${args.options.id}' from site '${args.options.webUrl}'...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/Web/Alerts/GetById('${formatting.encodeQueryParameter(args.options.id)}')?$expand=List,User,List/Rootfolder&$select=*,List/Id,List/Title,List/Rootfolder/ServerRelativeUrl`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const res = await request.get<any>(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoWebAlertGetCommand();
