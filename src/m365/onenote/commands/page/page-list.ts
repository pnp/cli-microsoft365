import { OnenotePage } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';
import { formatting } from '../../../../utils/formatting.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  userId: z.string().refine(id => validation.isValidGuid(id), {
    error: e => `'${e.input}' is not a valid GUID.`
  }).optional(),
  userName: z.string().optional(),
  groupId: z.string().refine(id => validation.isValidGuid(id), {
    error: e => `'${e.input}' is not a valid GUID.`
  }).optional(),
  groupName: z.string().optional(),
  webUrl: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
    error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
  }).optional().alias('u')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class OneNotePageListCommand extends GraphDelegatedCommand {
  public get name(): string {
    return commands.PAGE_LIST;
  }

  public get description(): string {
    return 'Retrieves a list of OneNote pages.';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => {
        const opts = [options.userId, options.userName, options.groupId, options.groupName, options.webUrl];
        const defined = opts.filter(item => item !== undefined);
        return defined.length <= 1;
      }, {
        error: 'Specify userId, userName, groupId, groupName, or webUrl, but not multiple.',
        params: {
          customCode: 'optionSet',
          options: ['userId', 'userName', 'groupId', 'groupName', 'webUrl']
        }
      });
  }

  public defaultProperties(): string[] | undefined {
    return ['createdDateTime', 'title', 'id'];
  }

  private async getEndpointUrl(args: CommandArgs): Promise<string> {
    let endpoint: string = `${this.resource}/v1.0/`;

    if (args.options.userId) {
      endpoint += `users/${args.options.userId}`;
    }
    else if (args.options.userName) {
      endpoint += `users/${formatting.encodeQueryParameter(args.options.userName)}`;
    }
    else if (args.options.groupId) {
      endpoint += `groups/${args.options.groupId}`;
    }
    else if (args.options.groupName) {
      const groupId = await entraGroup.getGroupIdByDisplayName(args.options.groupName);
      endpoint += `groups/${groupId}`;
    }
    else if (args.options.webUrl) {
      const siteId = await spo.getSpoGraphSiteId(args.options.webUrl);
      endpoint += `sites/${siteId}`;
    }
    else {
      endpoint += 'me';
    }
    endpoint += '/onenote/pages';
    return endpoint;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const endpoint = await this.getEndpointUrl(args);
      const items = await odata.getAllItems<OnenotePage>(endpoint);
      await logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new OneNotePageListCommand();