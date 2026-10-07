import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
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
  newDisplayName: z.string().optional(),
  description: z.string().optional(),
  privacy: z.string().optional()
});

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
      })
      .refine(opts => !opts.newDisplayName || opts.newDisplayName.length <= 255, {
        message: "The maximum amount of characters for 'newDisplayName' is 255.",
        params: { customCode: 'required' }
      })
      .refine(opts => !opts.description || opts.description.length <= 1024, {
        message: "The maximum amount of characters for 'description' is 1024.",
        params: { customCode: 'required' }
      })
      .refine(opts => {
        if (!opts.privacy) {
          return true;
        }
        const validPrivacy = ['public', 'private'];
        return validPrivacy.indexOf(opts.privacy.toLowerCase()) !== -1;
      }, {
        error: opts => `${opts.privacy} is not a valid privacy. Allowed values are public, private`,
        params: { customCode: 'required' }
      })
      .refine(opts => opts.newDisplayName || opts.description || opts.privacy, {
        message: 'Specify at least newDisplayName, description, or privacy.',
        params: { customCode: 'required' }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let communityId = args.options.id;

    if (args.options.displayName) {
      communityId = (await vivaEngage.getCommunityByDisplayName(args.options.displayName, ['id'])).id!;
    }
    else if (args.options.entraGroupId) {
      communityId = (await vivaEngage.getCommunityByEntraGroupId(args.options.entraGroupId, ['id'])).id!;
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