import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
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
    newDisplayName: z.string().optional().refine(value => !value || value.length <= 255, {
      message: 'The maximum amount of characters is 255.'
    }),
    description: z.string().optional().refine(value => !value || value.length <= 1024, {
      message: 'The maximum amount of characters is 1024.'
    }),
    privacy: z.enum(['public', 'private']).optional()
  })
  .strict();
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class VivaEngageCommunitySetCommand extends GraphCommand {
  public get name(): string {
    return commands.ENGAGE_COMMUNITY_SET;
  }

  public get description(): string {
    return 'Updates an existing Viva Engage community';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => Object.values([options.id, options.displayName, options.entraGroupId]).filter(x => typeof x !== 'undefined').length === 1, {
        message: `Specify either id, displayName, or entraGroupId, but not multiple.`
      })
      .refine(options => options.newDisplayName || options.description || options.privacy, {
        message: 'Specify at least newDisplayName, description, or privacy.'
      })
      .refine(options => (!options.id && !options.displayName && !options.entraGroupId) || options.id || options.displayName ||
        (options.entraGroupId && validation.isValidGuid(options.entraGroupId)), options => ({
        message: `The '${options.entraGroupId}' must be a valid GUID`,
        path: ['entraGroupId']
      }));
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {

    let communityId = args.options.id;

    if (args.options.displayName) {
      communityId = await vivaEngage.getCommunityIdByDisplayName(args.options.displayName);
    }

    if (args.options.entraGroupId) {
      communityId = await vivaEngage.getCommunityIdByEntraGroupId(args.options.entraGroupId);
    }

    if (this.verbose) {
      await logger.logToStderr(`Updating Viva Engage community with ID ${communityId}...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/employeeExperience/communities/${communityId}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: {
        description: args.options.description,
        displayName: args.options.newDisplayName,
        privacy: args.options.privacy
      }
    };

    try {
      await request.patch(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new VivaEngageCommunitySetCommand();