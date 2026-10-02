import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  teamId: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'teamId must be a valid GUID'
    })
    .alias('i'),
  allowUserEditMessages: z.boolean().optional(),
  allowUserDeleteMessages: z.boolean().optional(),
  allowOwnerDeleteMessages: z.boolean().optional(),
  allowTeamMentions: z.boolean().optional(),
  allowChannelMentions: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;
interface CommandArgs {
  options: Options;
}

class TeamsMessagingSettingsSetCommand extends GraphCommand {
  private static booleanProps: string[] = [
    'allowUserEditMessages',
    'allowUserDeleteMessages',
    'allowOwnerDeleteMessages',
    'allowTeamMentions',
    'allowChannelMentions'
  ];

  public get name(): string {
    return commands.MESSAGINGSETTINGS_SET;
  }

  public get description(): string {
    return 'Updates messaging settings of a Microsoft Teams team';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const data: any = {
      messagingSettings: {}
    };
    TeamsMessagingSettingsSetCommand.booleanProps.forEach((p: string) => {
      if (typeof (args.options as any)[p] !== 'undefined') {
        data.messagingSettings[p] = (args.options as any)[p];
      }
    });

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(args.options.teamId)}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      data: data,
      responseType: 'json'
    };

    try {
      await request.patch(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TeamsMessagingSettingsSetCommand();