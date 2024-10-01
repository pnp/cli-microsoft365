import SpoCommand from '../../../base/SpoCommand.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { zod } from '../../../../utils/zod.js';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import request, { CliRequestOptions } from '../../../../request.js';

const options = globalOptionsZod
  .extend({
    siteUrl: zod.alias('u', z.string()
      .refine(url => validation.isValidSharePointUrl(url) === true, url => ({
        message: `'${url}' is not a valid SharePoint Online site URL.`
      }))
    ),
    capability: z.enum(['full', 'limited', 'ownersOnly'])
  })
  .strict();
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoSiteSharingPermissionSetCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_SHARINGPERMISSION_SET;
  }

  public get description(): string {
    return 'Controls how a site and its components can be shared';
  }

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Updating sharing permissions for site '${args.options.siteUrl}'...`);
      }

      const { capability } = args.options;

      if (this.verbose) {
        await logger.logToStderr(`Updating site sharing permissions...`);
      }
      const requestOptionsWeb: CliRequestOptions = {
        url: `${args.options.siteUrl}/_api/Web`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json',
        data: {
          MembersCanShare: capability === 'full' || capability === 'limited'
        }
      };
      await request.patch(requestOptionsWeb);

      if (this.verbose) {
        await logger.logToStderr(`Updating associated member group sharing permissions...`);
      }

      const requestOptionsMemberGroup: CliRequestOptions = {
        url: `${args.options.siteUrl}/_api/Web/AssociatedMemberGroup`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json',
        data: {
          AllowMembersEditMembership: capability === 'full'
        }
      };

      await request.patch(requestOptionsMemberGroup);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteSharingPermissionSetCommand();