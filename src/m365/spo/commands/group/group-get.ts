import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import { zod } from '../../../../utils/zod.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

enum AssociatedGroup {
  Owner = 'Owner',
  Member = 'Member',
  Visitor = 'Visitor'
}

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  id: z.number().optional().alias('i'),
  name: z.string().optional(),
  associatedGroup: zod.coercedEnum(AssociatedGroup).optional()
});

export declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoGroupGetCommand extends SpoCommand {
  public get name(): string {
    return commands.GROUP_GET;
  }

  public get description(): string {
    return 'Gets site group';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(o => [o.id, o.name, o.associatedGroup].filter(x => x !== undefined).length === 1, {
        error: `Specify either 'id', 'name', or 'associatedGroup', but not multiple.`,
        params: {
          customCode: 'optionSet',
          options: ['id', 'name', 'associatedGroup']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving information for group in site at ${args.options.webUrl}...`);
    }

    let requestUrl: string = '';

    if (args.options.id) {
      requestUrl = `${args.options.webUrl}/_api/web/sitegroups/GetById('${args.options.id}')`;
    }
    else if (args.options.name) {
      requestUrl = `${args.options.webUrl}/_api/web/sitegroups/GetByName('${formatting.encodeQueryParameter(args.options.name as string)}')`;
    }
    else if (args.options.associatedGroup) {
      switch (args.options.associatedGroup.toLowerCase()) {
        case 'owner':
          requestUrl = `${args.options.webUrl}/_api/web/AssociatedOwnerGroup`;
          break;
        case 'member':
          requestUrl = `${args.options.webUrl}/_api/web/AssociatedMemberGroup`;
          break;
        case 'visitor':
          requestUrl = `${args.options.webUrl}/_api/web/AssociatedVisitorGroup`;
          break;
      }
    }

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      method: 'GET',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const groupInstance = await request.get(requestOptions);
      await logger.log(groupInstance);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoGroupGetCommand();
