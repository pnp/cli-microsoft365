import { Channel } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { teams } from '../../../../utils/teams.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string()
    .refine(val => validation.isValidTeamsChannelId(val), {
      message: 'The value is not a valid Teams channel id.'
    })
    .optional()
    .alias('i'),
  name: z.string()
    .refine(val => val.toLowerCase() !== 'general', {
      message: 'General channel cannot be updated.'
    })
    .optional()
    .alias('n'),
  description: z.string()
    .optional(),
  newName: z.string()
    .optional(),
  teamId: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value must be a valid GUID.'
    })
    .optional(),
  teamName: z.string()
    .optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsChannelSetCommand extends GraphCommand {
  public get name(): string {
    return commands.CHANNEL_SET;
  }
  public get description(): string {
    return 'Updates properties of the specified channel in the given Microsoft Teams team';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.id, opts.name].filter(x => x !== undefined).length === 1, {
        message: 'Specify either id or name, but not both.',
        params: {
          customCode: 'optionSet',
          options: ['id', 'name']
        }
      })
      .refine(opts => [opts.teamId, opts.teamName].filter(x => x !== undefined).length === 1, {
        message: 'Specify either teamId or teamName, but not both.',
        params: {
          customCode: 'optionSet',
          options: ['teamId', 'teamName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const teamId = await this.getTeamId(args);
      const channelId: string = await this.getChannelId(teamId, args);

      const data: any = this.mapRequestBody(args.options);
      const requestOptionsPatch: any = {
        url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(teamId)}/channels/${formatting.encodeQueryParameter(channelId)}`,
        headers: {
          'accept': 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: data
      };

      await request.patch(requestOptionsPatch);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private mapRequestBody(options: Options): any {
    const requestBody: any = {};

    if (options.newName) {
      requestBody.displayName = options.newName;
    }

    if (options.description) {
      requestBody.description = options.description;
    }

    return requestBody;
  }

  private async getTeamId(args: CommandArgs): Promise<string> {
    if (args.options.teamId) {
      return args.options.teamId;
    }

    return teams.getTeamIdByDisplayName(args.options.teamName!);
  }

  private async getChannelId(teamId: string, args: CommandArgs): Promise<string> {
    if (args.options.id) {
      return args.options.id;
    }

    const channelRequestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(teamId)}/channels?$filter=displayName eq '${formatting.encodeQueryParameter(args.options.name!)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const res: { value: Channel[] } = await request.get<{ value: Channel[] }>(channelRequestOptions);
    const channelItem: Channel | undefined = res.value[0];

    if (!channelItem) {
      throw `The specified channel does not exist in this Microsoft Teams team`;
    }

    return channelItem.id!;
  }
}

export default new TeamsChannelSetCommand();