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
  name: z.string().alias('n'),
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
  }).alias('d')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoAppPageSetCommand extends SpoCommand {
  public get name(): string {
    return commands.APPPAGE_SET;
  }

  public get description(): string {
    return 'Updates the single-part app page';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${args.options.webUrl}/_api/sitepages/Pages/UpdateFullPageApp`,
      headers: {
        'content-type': 'application/json;odata=nometadata',
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json',
      data: {
        serverRelativeUrl: `${urlUtil.getServerRelativePath(args.options.webUrl, 'SitePages')}/${args.options.name}`,
        webPartDataAsJson: args.options.webPartData
      }
    };

    try {
      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}
export default new SpoAppPageSetCommand();
