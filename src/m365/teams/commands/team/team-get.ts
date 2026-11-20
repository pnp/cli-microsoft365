import { Team } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { teams } from '../../../../utils/teams.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value must be a valid GUID.'
    })
    .optional()
    .alias('i'),
  name: z.string().optional().alias('n')
});

export declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsTeamGetCommand extends GraphCommand {
  public get name(): string {
    return commands.TEAM_GET;
  }

  public get description(): string {
    return 'Gets information about the specified Microsoft Teams team';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema.refine(opts => [opts.id, opts.name].filter(o => o !== undefined).length === 1, {
      message: `Specify either 'id' or 'name', but not both.`,
      params: {
        customCode: 'optionSet',
        options: ['id', 'name']
      }
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let team: Team;
      if (args.options.id) {
        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(args.options.id)}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };
        team = await request.get<Team>(requestOptions);
      }
      else {
        team = await teams.getTeamByDisplayName(args.options.name!);
      }

      await logger.log(team);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TeamsTeamGetCommand();