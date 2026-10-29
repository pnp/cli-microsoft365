import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
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
  id: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value must be a valid GUID.'
    })
    .optional()
    .alias('i'),
  name: z.string().optional().alias('n'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsTeamRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.TEAM_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified Microsoft Teams team';
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

  private async getTeamId(args: CommandArgs): Promise<string> {
    if (args.options.id) {
      return args.options.id;
    }

    return teams.getTeamIdByDisplayName(args.options.name!);
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeTeam = async (): Promise<void> => {
      try {
        const teamId: string = await this.getTeamId(args);
        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/groups/${formatting.encodeQueryParameter(teamId)}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeTeam();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the team ${args.options.id ? args.options.id : args.options.name}?` });

      if (result) {
        await removeTeam();
      }
    }
  }
}

export default new TeamsTeamRemoveCommand();