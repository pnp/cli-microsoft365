import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { SiteDesignRun } from './SiteDesignRun.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string().refine(val => validation.isValidSharePointUrl(val) === true, { message: 'The value is not a valid SharePoint Online site URL.' }).alias('u'),
  siteDesignId: z.string().refine(val => validation.isValidGuid(val), { message: 'The value must be a valid GUID.' }).optional().alias('i')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoSiteDesignRunListCommand extends SpoCommand {
  public get name(): string {
    return commands.SITEDESIGN_RUN_LIST;
  }

  public get description(): string {
    return 'Lists information about site designs applied to the specified site';
  }

  public defaultProperties(): string[] | undefined {
    return ['ID', 'SiteDesignID', 'SiteDesignTitle', 'StartTime'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const data: any = {};
    if (args.options.siteDesignId) {
      data.siteDesignId = args.options.siteDesignId;
    }

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRun`,
      headers: {
        accept: 'application/json;odata=nometadata',
        'content-type': 'application/json;odata=nometadata'
      },
      data: data,
      responseType: 'json'
    };

    try {
      const res: { value: SiteDesignRun[] } = await request.post<{ value: SiteDesignRun[] }>(requestOptions);
      if (args.options.output !== 'json') {
        res.value.forEach(d => {
          d.StartTime = new Date(parseInt(d.StartTime)).toLocaleString();
        });
      }

      await logger.log(res.value);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteDesignRunListCommand();