import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import { zod } from '../../../../utils/zod.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

enum CheckinType {
  Minor = 0,
  Major = 1,
  Overwrite = 2
}

enum CheckinTypeOption {
  Minor = 'Minor',
  Major = 'Major',
  Overwrite = 'Overwrite'
}

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  url: z.string().optional(),
  id: z.uuid().optional().alias('i'),
  type: zod.coercedEnum(CheckinTypeOption).optional().alias('t'),
  comment: z.string()
    .max(1023, { error: 'The length of the comment must be less than 1024 letters' })
    .optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFileCheckinCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_CHECKIN;
  }

  public get description(): string {
    return 'Checks in specified file';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.url, opts.id].filter(o => o !== undefined).length === 1, {
        error: `Specify 'url' or 'id', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['url', 'id']
        }
      });
  }

  protected getExcludedOptionsWithUrls(): string[] | undefined {
    return ['url'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let type: CheckinType = CheckinType.Major;
    if (args.options.type) {
      switch (args.options.type.toLowerCase()) {
        case 'minor':
          type = CheckinType.Minor;
          break;
        case 'overwrite':
          type = CheckinType.Overwrite;
      }
    }

    let comment: string = '';
    if (args.options.comment) {
      comment = formatting.encodeQueryParameter(args.options.comment);
    }

    let requestUrl: string = '';
    if (args.options.id) {
      requestUrl = `${args.options.webUrl}/_api/web/GetFileById('${formatting.encodeQueryParameter(args.options.id)}')/checkin(comment='${comment}',checkintype=${type})`;
    }

    if (args.options.url) {
      const serverRelativePath = urlUtil.getServerRelativePath(args.options.webUrl, args.options.url);
      requestUrl = `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')/checkin(comment='${comment}',checkintype=${type})`;
    }

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoFileCheckinCommand();