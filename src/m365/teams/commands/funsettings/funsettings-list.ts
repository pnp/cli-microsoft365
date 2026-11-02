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
    .alias('i')
});

declare type Options = z.infer<typeof options>;
interface CommandArgs {
  options: Options;
}

class TeamsFunSettingsListCommand extends GraphCommand {
  public get name(): string {
    return commands.FUNSETTINGS_LIST;
  }

  public get description(): string {
    return 'Lists fun settings for the specified Microsoft Teams team';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Listing fun settings for team ${args.options.teamId}`);
    }
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(args.options.teamId)}?$select=funSettings`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      const res: { funSettings: any } = await request.get<{ funSettings: any }>(requestOptions);
      await logger.log(res.funSettings);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TeamsFunSettingsListCommand();