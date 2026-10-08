import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import { accessToken } from '../../../../utils/accessToken.js';
import auth from '../../../../Auth.js';
import { formatting } from '../../../../utils/formatting.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { setTimeout } from 'timers/promises';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  displayName: z.string(),
  description: z.string(),
  privacy: z.enum(['public', 'private']),
  adminEntraIds: z.string().optional(),
  adminEntraUserNames: z.string().optional(),
  wait: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class VivaEngageCommunityAddCommand extends GraphCommand {
  private pollingInterval: number = 5000;

  public get name(): string {
    return commands.ENGAGE_COMMUNITY_ADD;
  }

  public get description(): string {
    return 'Creates a new community in Viva Engage';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => opts.displayName.length <= 255, {
        message: "The maximum amount of characters for 'displayName' is 255.",
        params: { customCode: 'required' }
      })
      .refine(opts => opts.description.length <= 1024, {
        message: "The maximum amount of characters for 'description' is 1024.",
        params: { customCode: 'required' }
      })
      .superRefine((opts, ctx) => {
        if (opts.adminEntraIds) {
          const items = opts.adminEntraIds.split(',').map(s => s.trim());
          const invalid = items.filter(item => !validation.isValidGuid(item));
          if (invalid.length > 0) {
            ctx.addIssue({
              code: z.ZodIssueCode.custom,
              message: `The following GUIDs are invalid for the option 'adminEntraIds': ${invalid.join(', ')}`,
              params: { customCode: 'required' }
            });
          }
        }
      })
      .refine(opts => {
        if (!opts.adminEntraIds) {
          return true;
        }
        return formatting.splitAndTrim(opts.adminEntraIds).length <= 20;
      }, {
        message: 'Maximum of 20 admins allowed. Please reduce the number of users and try again.',
        params: { customCode: 'required' }
      })
      .superRefine((opts, ctx) => {
        if (opts.adminEntraUserNames) {
          const items = opts.adminEntraUserNames.split(',').map(s => s.trim());
          const invalid = items.filter(item => !validation.isValidUserPrincipalName(item));
          if (invalid.length > 0) {
            ctx.addIssue({
              code: z.ZodIssueCode.custom,
              message: `The following user principal names are invalid for the option 'adminEntraUserNames': ${invalid.join(', ')}`,
              params: { customCode: 'required' }
            });
          }
        }
      })
      .refine(opts => {
        if (!opts.adminEntraUserNames) {
          return true;
        }
        return formatting.splitAndTrim(opts.adminEntraUserNames).length <= 20;
      }, {
        message: 'Maximum of 20 admins allowed. Please reduce the number of users and try again.',
        params: { customCode: 'required' }
      })
      .refine(opts => !opts.adminEntraIds || !opts.adminEntraUserNames, {
        message: 'Specify either adminEntraIds or adminEntraUserNames, but not both.',
        params: {
          customCode: 'optionSet',
          options: ['adminEntraIds', 'adminEntraUserNames']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const { displayName, description, privacy, adminEntraIds, adminEntraUserNames, wait } = args.options;

    const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken);
    if (isAppOnlyAccessToken && !adminEntraIds && !adminEntraUserNames) {
      this.handleError(`Specify at least one admin using either adminEntraIds or adminEntraUserNames options when using application permissions.`);
    }

    if (this.verbose) {
      await logger.logToStderr(`Creating a Viva Engage community with display name '${displayName}'...`);
    }

    try {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/beta/employeeExperience/communities`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json'
        },
        responseType: 'json',
        fullResponse: true,
        data: {
          displayName: displayName,
          description: description,
          privacy: privacy
        }
      };

      const entraIds = await this.getGraphUserUrls(args.options);
      if (entraIds.length > 0) {
        requestOptions.data['owners@odata.bind'] = entraIds;
      }

      const res = await request.post<{ headers: { location: string } }>(requestOptions);

      const location = res.headers.location;

      if (!wait) {
        await logger.log(location);
        return;
      }

      let status: string;
      do {
        if (this.verbose) {
          await logger.logToStderr(`Community still provisioning. Retrying in ${this.pollingInterval / 1000} seconds...`);
        }

        await setTimeout(this.pollingInterval);

        if (this.verbose) {
          await logger.logToStderr(`Checking create community operation status...`);
        }

        const operation = await request.get<any>({
          url: location,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        });
        status = operation.status;

        if (this.verbose) {
          await logger.logToStderr(`Community creation operation status: ${status}`);
        }

        if (status === 'failed') {
          throw `Community creation failed: ${operation.statusDetail}`;
        }

        if (status === 'succeeded') {
          await logger.log(operation);
        }
      }
      while (status === 'notStarted' || status === 'running');
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getGraphUserUrls(options: Options): Promise<string[]> {
    let entraIds: string[] = [];

    if (options.adminEntraIds) {
      entraIds = formatting.splitAndTrim(options.adminEntraIds);
    }
    else if (options.adminEntraUserNames) {
      entraIds = await entraUser.getUserIdsByUpns(formatting.splitAndTrim(options.adminEntraUserNames));
    }

    const graphUserUrls = entraIds.map(id => `${this.resource}/beta/users/${id}`);
    return graphUserUrls;
  }
}

export default new VivaEngageCommunityAddCommand();