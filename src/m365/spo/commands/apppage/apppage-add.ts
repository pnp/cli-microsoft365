import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string().alias('u'),
  title: z.string().alias('t'),
  webPartData: z.string().refine(val => {
    try {
      JSON.parse(val);
      return true;
    }
    catch {
      return false;
    }
  }, {
    error: 'Specified webPartData is not a valid JSON string.'
  }).alias('d'),
  addToQuickLaunch: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoAppPageAddCommand extends SpoCommand {
  public get name(): string {
    return commands.APPPAGE_ADD;
  }

  public get description(): string {
    return 'Creates a single-part app page';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const createPageRequestOptions: CliRequestOptions = {
      url: `${args.options.webUrl}/_api/sitepages/Pages/CreateAppPage`,
      headers: {
        'content-type': 'application/json;odata=nometadata',
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json',
      data: {
        webPartDataAsJson: args.options.webPartData
      }
    };

    try {
      const page = await request.post<{ value: string }>(createPageRequestOptions);

      const pageUrl: string = page.value;

      let requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${urlUtil.getServerRelativeSiteUrl(args.options.webUrl)}/${pageUrl}')?$expand=ListItemAllFields`,
        headers: {
          'content-type': 'application/json;charset=utf-8',
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const file = await request.get<{ ListItemAllFields: { Id: string; }; }>(requestOptions);

      requestOptions = {
        url: `${args.options.webUrl}/_api/sitepages/Pages/UpdateAppPage`,
        headers: {
          'content-type': 'application/json;odata=nometadata',
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json',
        data: {
          pageId: file.ListItemAllFields.Id,
          webPartDataAsJson: args.options.webPartData,
          title: args.options.title,
          includeInNavigation: args.options.addToQuickLaunch
        }
      };

      const res = await request.post(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoAppPageAddCommand();
