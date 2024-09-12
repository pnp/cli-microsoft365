import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import { cli } from '../../../../cli/cli.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { vivaEngage } from '../../../../utils/vivaEngage.js';
import { zod } from '../../../../utils/zod.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';

const options = globalOptionsZod
  .extend({
    id: zod.alias('i', z.string().optional()),
    displayName: zod.alias('d', z.string().optional()),
    entraGroupId: z.string().optional(),
    force: z.boolean().optional()
  })
  .strict();
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class VivaEngageCommunityRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.ENGAGE_COMMUNITY_REMOVE;
  }

  public get description(): string {
    return 'Removes a community';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => Object.values([options.id, options.displayName, options.entraGroupId]).filter(x => typeof x !== 'undefined').length === 1, {
        message: `Specify either id, displayName, or entraGroupId, but not multiple.`
      })
      .refine(options => (!options.id && !options.displayName && !options.entraGroupId) || options.id || options.displayName ||
        (options.entraGroupId && validation.isValidGuid(options.entraGroupId)), options => ({
        message: `The '${options.entraGroupId}' must be a valid GUID`,
        path: ['entraGroupId']
      }));
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeCommunity = async (): Promise<void> => {
      try {
        let communityId = args.options.id;

        if (args.options.displayName) {
          communityId = await vivaEngage.getCommunityIdByDisplayName(args.options.displayName);
        }
        else if (args.options.entraGroupId) {
          communityId = await vivaEngage.getCommunityIdByEntraGroupId(args.options.entraGroupId);
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
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove Viva Engage community '${args.options.id || args.options.displayName}'?` });

      if (result) {
        await removeCommunity();
      }
    }
  }
}

export default new VivaEngageCommunityRemoveCommand();