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
    disabled: z.boolean().optional(),
    ownerGroup: z.boolean().optional(),
    email: z.string()
      .refine(email => validation.isValidUserPrincipalName(email), email => ({
        message: `'${email}' is not a valid email address.`
      })).optional(),
    message: z.string().optional()
  })
  .strict();
type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoSiteAccessRequestSettingSetCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_ACCESSREQUEST_SETTING_SET;
  }

  public get description(): string {
    return 'Update access requests for a specific site';
  }

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(o => [o.disabled, o.ownerGroup, o.email].filter(v => v !== undefined).length === 1, {
        message: 'Specify exactly one of disabled, ownerGroup, or email'
      })
      .refine(o => !(o.disabled && typeof o.message !== 'undefined'), {
        message: 'The message option cannot be used when disabled is specified'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Updating access requests for site '${args.options.siteUrl}'...`);
    }
    try {
      const { siteUrl, ownerGroup, email, message } = args.options;

      const requestAccessEmail: string = email || '';

      if (this.verbose) {
        await logger.logToStderr(`Updating RequestAccessEmail to '${requestAccessEmail}'...`);
      }

      const requestPatchWeb: CliRequestOptions = {
        url: `${siteUrl}/_api/Web`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json',
        data: {
          RequestAccessEmail: requestAccessEmail
        }
      };
      await request.patch(requestPatchWeb);

      const useAccessRequestDefault = !!ownerGroup;

      if (this.verbose) {
        await logger.logToStderr(`Updating UseAccessRequestDefault to '${useAccessRequestDefault}'...`);
      }

      const requestUseDefault: CliRequestOptions = {
        url: `${siteUrl}/_api/Web/SetUseAccessRequestDefaultAndUpdate`,
        headers: {
          accept: 'application/json;odata=nometadata',
          'content-type': 'application/json;odata=nometadata'
        },
        responseType: 'json',
        data: {
          useAccessRequestDefault: useAccessRequestDefault
        }
      };
      await request.post(requestUseDefault);

      if (message !== undefined) {

        if (this.verbose) {
          await logger.logToStderr(`Updating access request message to '${message}'...`);
        }

        const requestSetMessage: CliRequestOptions = {
          url: `${siteUrl}/_api/Web/SetAccessRequestSiteDescriptionAndUpdate`,
          headers: {
            accept: 'application/json;odata=nometadata',
            'content-type': 'application/json;odata=nometadata'
          },
          responseType: 'json',
          data: {
            description: message
          }
        };
        await request.post(requestSetMessage);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteAccessRequestSettingSetCommand();


