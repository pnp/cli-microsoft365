import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import Command, { globalOptionsZod } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import teamGetCommand, { Options as TeamsTeamGetOptions } from './team-get.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  teamId: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value must be a valid GUID.'
    })
    .optional()
    .alias('i'),
  teamName: z.string().optional().alias('n')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsTeamAppListCommand extends GraphCommand {
  public get name(): string {
    return commands.TEAM_APP_LIST;
  }

  public get description(): string {
    return 'Lists apps installed in the specified team';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'distributionMethod'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema.refine(opts => [opts.teamId, opts.teamName].filter(o => o !== undefined).length === 1, {
      message: `Specify either 'teamId' or 'teamName', but not both.`,
      params: {
        customCode: 'optionSet',
        options: ['teamId', 'teamName']
      }
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving installed apps for team '${args.options.teamId || args.options.teamName}'`);
      }

      const teamId: string = await this.getTeamId(args);
      const res = await odata.getAllItems<any>(`${this.resource}/v1.0/teams/${teamId}/installedApps?$expand=teamsApp,teamsAppDefinition`);

      if (!cli.shouldTrimOutput(args.options.output)) {
        await logger.log(res);
      }
      else {
        //converted to text friendly output
        await logger.log(res.map(i => {
          return {
            id: i.id,
            displayName: i.teamsApp.displayName,
            distributionMethod: i.teamsApp.distributionMethod
          };
        }));
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getTeamId(args: CommandArgs): Promise<string> {
    if (args.options.teamId) {
      return args.options.teamId;
    }

    const teamGetOptions: TeamsTeamGetOptions = {
      name: args.options.teamName,
      debug: this.debug,
      verbose: this.verbose
    };

    const commandOutput = await cli.executeCommandWithOutput(teamGetCommand as Command, { options: { ...teamGetOptions, _: [] } });
    const team = JSON.parse(commandOutput.stdout);
    return team.id;
  }
}

export default new TeamsTeamAppListCommand();