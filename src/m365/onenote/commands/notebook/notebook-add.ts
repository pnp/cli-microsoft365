import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import { spo } from '../../../../utils/spo.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';
import { formatting } from '../../../../utils/formatting.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  name: z.string()
    .refine(name => name.length <= 128, {
      error: 'The specified name is too long. It should be at most 128 characters.'
    })
    .refine(name => !/[?*/:<>|'"]/.test(name), {
      error: `The specified name contains invalid characters. It cannot contain ?*/:<>|'". Please remove them and try again.`
    }).alias('n'),
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

class OneNoteNotebookAddCommand extends GraphDelegatedCommand {
  public get name(): string {
    return commands.NOTEBOOK_ADD;
  }

  public get description(): string {
    return 'Create a new OneNote notebook';
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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Creating OneNote notebook ${args.options.name}`);
      }

      const requestUrl = await this.getRequestUrl(args);
      const requestOptions: CliRequestOptions = {
        url: requestUrl,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': "application/json"
        },
        responseType: 'json',
        data: {
          displayName: args.options.name
        }
      };

      const response = await request.post(requestOptions);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getRequestUrl(args: CommandArgs): Promise<string> {
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
    endpoint += '/onenote/notebooks';
    return endpoint;
  }
}

export default new OneNoteNotebookAddCommand();