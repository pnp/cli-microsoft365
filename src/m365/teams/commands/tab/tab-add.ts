import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.object({
  ...globalOptionsZod.shape,
  teamId: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value must be a valid GUID.'
    })
    .alias('i'),
  channelId: z.string()
    .refine(val => validation.isValidTeamsChannelId(val), {
      message: 'The value must be a valid Teams channel ID.'
    })
    .alias('c'),
  appId: z.string(),
  appName: z.string(),
  contentUrl: z.string(),
  entityId: z.string().optional(),
  removeUrl: z.string().optional(),
  websiteUrl: z.string().optional()
}).catchall(z.unknown());

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsTabAddCommand extends GraphCommand {
  public get name(): string {
    return commands.TAB_ADD;
  }
  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public get description(): string {
    return 'Adds a tab to the specified channel';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const data: any = this.mapRequestBody(args.options);
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(args.options.teamId)}/channels/${args.options.channelId}/tabs`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json;odata=nometadata'
      },
      data: data,
      responseType: 'json'
    };

    try {
      const res = await request.post(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private mapRequestBody(options: Options): any {
    const requestBody: any = {};
    requestBody['configuration'] = {};
    const excludeOptions: string[] = [
      'debug',
      'verbose',
      'teamId',
      'channelId',
      'appId',
      'appName',
      'entityId',
      'contentUrl',
      'removeUrl',
      'websiteUrl',
      'output'
    ];
    if (options.appName) {
      requestBody.displayName = options.appName;
    }
    if (options.appId) {
      requestBody['teamsApp@odata.bind'] = `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/${options.appId}`;
    }
    if (options.contentUrl) {
      requestBody.configuration.contentUrl = options.contentUrl;
    }
    if (options.entityId) {
      requestBody.configuration.entityId = options.entityId;
    }
    if (options.removeUrl) {
      requestBody.configuration.removeUrl = options.removeUrl;
    }
    if (options.websiteUrl) {
      requestBody.configuration.websiteUrl = options.websiteUrl;
    }
    Object.keys(options).forEach(key => {
      if (excludeOptions.indexOf(key) === -1) {
        requestBody.configuration[key] = `${(<any>options)[key]}`;
      }
    });
    return requestBody;
  }
}

export default new TeamsTabAddCommand();