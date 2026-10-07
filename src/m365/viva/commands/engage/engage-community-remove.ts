import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { cli } from '../../../../cli/cli.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import { vivaEngage } from '../../../../utils/vivaEngage.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().optional(),
  displayName: z.string().optional(),
  entraGroupId: z.string().optional(),
  force: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class VivaEngageCommunityRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.ENGAGE_COMMUNITY_REMOVE;
  }
  public get description(): string {
    return 'Removes a Viva Engage community';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.id, opts.displayName, opts.entraGroupId].filter(x => x !== undefined).length === 1, {
        message: 'Specify either id, displayName, or entraGroupId, but not multiple.',
        params: {
          customCode: 'optionSet',
          options: ['id', 'displayName', 'entraGroupId']
        }
      })
      .refine(opts => !opts.entraGroupId || validation.isValidGuid(opts.entraGroupId), {
        error: opts => `${opts.entraGroupId} is not a valid GUID for the option 'entraGroupId'.`,
        params: { customCode: 'required' }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {

    const removeCommunity = async (): Promise<void> => {
      try {
        let communityId = args.options.id;

        if (args.options.displayName) {
          communityId = (await vivaEngage.getCommunityByDisplayName(args.options.displayName, ['id'])).id;
        }
        else if (args.options.entraGroupId) {
          communityId = (await vivaEngage.getCommunityByEntraGroupId(args.options.entraGroupId, ['id'])).id;
        }

        if (args.options.verbose) {
          await logger.logToStderr(`Removing Viva Engage community with ID ${communityId}...`);
        }

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/employeeExperience/communities/${communityId}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          }
        };

        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeCommunity();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove Viva Engage community '${args.options.id || args.options.displayName || args.options.entraGroupId}'?` });

      if (result) {
        await removeCommunity();
      }
    }
  }
}

export default new VivaEngageCommunityRemoveCommand();