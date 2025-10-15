import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import { zod } from '../../../../utils/zod.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

const options = globalOptionsZod
  .extend({
    webUrl: zod.alias('u', z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
      message: 'Specify a valid SharePoint site URL'
    })),
    id: z.string().refine(id => validation.isValidGuid(id) === true, {
      message: 'Specify a valid GUID'
    })
  })
  .strict();
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoSiteAlertGetCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_ALERT_GET;
  }

  public get description(): string {
    return 'Retrieves details of a specific alert from a SharePoint site list';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${args.options.webUrl}/_api/Web/Alerts/GetById('${args.options.id}')?$expand=List,User,List/Rootfolder&$select=*,List/Id,List/Title,List/Rootfolder/ServerRelativeUrl`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const res = await request.get(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteAlertGetCommand();

