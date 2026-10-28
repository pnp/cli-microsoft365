import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import { zod } from '../../../../utils/zod.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

const VisibilityEnum = {
  Private: 'Private',
  Public: 'Public'
} as const;

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value must be a valid GUID.'
    })
    .alias('i'),
  name: z.string().alias('n'),
  partsToClone: z.string()
    .refine(val => val.replace(/\s/g, '').split(',').every(part => ['apps', 'channels', 'members', 'settings', 'tabs'].includes(part.toLowerCase())), {
      message: `The value must be a comma-separated list of the following values: apps, channels, members, settings, tabs.`
    })
    .alias('p'),
  description: z.string().optional().alias('d'),
  classification: z.string().optional().alias('c'),
  visibility: zod.coercedEnum(VisibilityEnum).optional().alias('v')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsTeamCloneCommand extends GraphCommand {
  public get name(): string {
    return commands.TEAM_CLONE;
  }

  public get description(): string {
    return 'Creates a clone of a Microsoft Teams team';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const data: any = {
      displayName: args.options.name,
      mailNickname: this.generateMailNickname(args.options.name),
      partsToClone: args.options.partsToClone
    };

    if (args.options.description) {
      data.description = args.options.description;
    }

    if (args.options.classification) {
      data.classification = args.options.classification;
    }

    if (args.options.visibility) {
      data.visibility = args.options.visibility;
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(args.options.id)}/clone`,
      headers: {
        'content-type': 'application/json',
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: data
    };

    try {
      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  /**
   * There is a know issue that the mailNickname is currently ignored and cannot be set by the user
   * However the mailNickname is still required by the payload so to deliver better user experience
   * the CLI generates mailNickname for the user 
   * so the user does not have to specify something that will be ignored.
   * For more see: https://docs.microsoft.com/en-us/graph/api/team-clone?view=graph-rest-1.0#request-data
   * This method has to be removed once the graph team fixes the issue and then the actual value
   * of the mailNickname would have to be specified by the CLI user.
   * @param displayName teams display name
   */
  private generateMailNickname(displayName: string): string {
    // currently the Microsoft Graph generates mailNickname in a similar fashion
    return `${displayName.replace(/[^a-zA-Z0-9]/g, "")}${Math.floor(Math.random() * 9999)}`;
  }
}

export default new TeamsTeamCloneCommand();