import { Channel, TeamsTab } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { teams } from '../../../../utils/teams.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  teamId: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value must be a valid GUID.'
    })
    .optional(),
  teamName: z.string().optional(),
  channelId: z.string()
    .refine(val => validation.isValidTeamsChannelId(val), {
      message: 'The value must be a valid Teams channel ID.'
    })
    .optional(),
  channelName: z.string().optional(),
  id: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value must be a valid GUID.'
    })
    .optional()
    .alias('i'),
  name: z.string().optional().alias('n')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsTabGetCommand extends GraphCommand {
  private teamId: string = "";
  private channelId: string = "";

  public get name(): string {
    return commands.TAB_GET;
  }

  public get description(): string {
    return 'Gets information about the specified Microsoft Teams tab';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.teamId, opts.teamName].filter(x => x !== undefined).length === 1, {
        message: `Specify either 'teamId' or 'teamName', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['teamId', 'teamName']
        }
      })
      .refine(opts => [opts.channelId, opts.channelName].filter(x => x !== undefined).length === 1, {
        message: `Specify either 'channelId' or 'channelName', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['channelId', 'channelName']
        }
      })
      .refine(opts => [opts.id, opts.name].filter(x => x !== undefined).length === 1, {
        message: `Specify either 'id' or 'name', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['id', 'name']
        }
      });
  }

  private async getTeamId(args: CommandArgs): Promise<string> {
    if (args.options.teamId) {
      return args.options.teamId;
    }

    return teams.getTeamIdByDisplayName(args.options.teamName!);
  }

  private async getChannelId(args: CommandArgs): Promise<string> {
    if (args.options.channelId) {
      return args.options.channelId;
    }

    const channelRequestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(this.teamId)}/channels?$filter=displayName eq '${formatting.encodeQueryParameter(args.options.channelName as string)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: Channel[] }>(channelRequestOptions);
    const channelItem: Channel | undefined = response.value[0];

    if (!channelItem) {
      throw 'The specified channel does not exist in the Microsoft Teams team';
    }

    return channelItem.id!;
  }

  private async getTabId(args: CommandArgs): Promise<string> {
    if (args.options.id) {
      return args.options.id;
    }

    const tabRequestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(this.teamId)}/channels/${formatting.encodeQueryParameter(this.channelId)}/tabs?$filter=displayName eq '${formatting.encodeQueryParameter(args.options.name as string)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: TeamsTab[] }>(tabRequestOptions);
    const tabItem: TeamsTab | undefined = response.value[0];

    if (!tabItem) {
      throw 'The specified tab does not exist in the Microsoft Teams team channel';
    }

    return tabItem.id!;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      this.teamId = await this.getTeamId(args);
      this.channelId = await this.getChannelId(args);
      const tabId: string = await this.getTabId(args);
      const endpoint: string = `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(this.teamId)}/channels/${formatting.encodeQueryParameter(this.channelId)}/tabs/${formatting.encodeQueryParameter(tabId)}?$expand=teamsApp`;

      const requestOptions: CliRequestOptions = {
        url: endpoint,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const res: TeamsTab = await request.get<TeamsTab>(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TeamsTabGetCommand();