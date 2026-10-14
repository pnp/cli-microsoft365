import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import { zod } from '../../../../utils/zod.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

enum RecycleBinItemType {
  files = 'files',
  listItems = 'listItems',
  folders = 'folders'
}

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  siteUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  type: zod.coercedEnum(RecycleBinItemType).optional().alias('t'),
  secondary: z.boolean().optional().alias('s')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoSiteRecycleBinItemListCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_RECYCLEBINITEM_LIST;
  }

  public get description(): string {
    return 'Lists items from recycle bin';
  }

  public defaultProperties(): string[] | undefined {
    return ['Id', 'Title', 'DirName'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving all items from recycle bin at ${args.options.siteUrl}...`);
    }

    const state: string = args.options.secondary ? '2' : '1';

    let requestUrl: string = `${args.options.siteUrl}/_api/site/RecycleBin?$filter=(ItemState eq ${state})`;

    if (typeof args.options.type !== 'undefined') {
      const type = SpoSiteRecycleBinItemListCommand.recycleBinItemType.find(item => item.value === args.options.type);
      if (typeof type !== 'undefined') {
        requestUrl += ` and (ItemType eq ${type.id})`;
      }
    }

    const requestOptions: any = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const response = await request.get<{ value: any[] }>(requestOptions);
      await logger.log(response.value);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private static recycleBinItemType: { id: number, value: string }[] = [
    { id: 1, value: 'files' },
    { id: 3, value: 'listItems' },
    { id: 5, value: 'folders' }
  ];
}

export default new SpoSiteRecycleBinItemListCommand();