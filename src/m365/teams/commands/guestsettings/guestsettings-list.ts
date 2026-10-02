import { Team } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from "../../../base/GraphCommand.js";
import commands from '../../commands.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  teamId: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'teamId must be a valid GUID'
    })
    .alias('i')
});

declare type Options = z.infer<typeof options>;
interface CommandArgs {
  options: Options;
}

class TeamsGuestSettingsListCommand extends GraphCommand {
  public get name(): string {
    return commands.GUESTSETTINGS_LIST;
  }

  public get description(): string {
    return 'Lists guest settings for a Microsoft Teams team';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(args.options.teamId)}?$select=guestSettings`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      const res: Team = await request.get<Team>(requestOptions);
      await logger.log(res.guestSettings);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TeamsGuestSettingsListCommand();