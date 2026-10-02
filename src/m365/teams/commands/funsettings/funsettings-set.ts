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
  allowGiphy: z.boolean().optional(),
  giphyContentRating: z.string()
    .refine(val => ['strict', 'moderate'].includes(val.toLowerCase()), {
      message: `giphyContentRating value must be 'Strict' or 'Moderate'`
    })
    .optional(),
  allowStickersAndMemes: z.boolean().optional(),
  allowCustomMemes: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;
interface CommandArgs {
  options: Options;
}

class TeamsFunSettingsSetCommand extends GraphCommand {
  private static booleanProps: string[] = [
    'allowGiphy',
    'allowStickersAndMemes',
    'allowCustomMemes'
  ];

  public get name(): string {
    return commands.FUNSETTINGS_SET;
  }

  public get description(): string {
    return 'Updates fun settings of a Microsoft Teams team';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Updating fun settings for team ${args.options.teamId}`);
      }

      const data: any = {
        funSettings: {}
      };
      TeamsFunSettingsSetCommand.booleanProps.forEach(p => {
        if (typeof (args.options as any)[p] !== 'undefined') {
          data.funSettings[p] = (args.options as any)[p];
        }
      });

      if (args.options.giphyContentRating) {
        data.funSettings.giphyContentRating = args.options.giphyContentRating;
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(args.options.teamId)}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        data: data,
        responseType: 'json'
      };

      await request.patch(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TeamsFunSettingsSetCommand();