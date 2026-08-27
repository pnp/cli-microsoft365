import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.looseObject({
  ...globalOptionsZod.shape,
  description: z.string().optional().alias('d'),
  headerEmphasis: z.enum(['0', '1', '2', '3'], {
    error: e => `${e.input} is not a valid value for headerEmphasis. Allowed values are 0|1|2|3`
  }).optional(),
  headerLayout: z.enum(['standard', 'compact']).optional(),
  megaMenuEnabled: z.boolean().optional(),
  quickLaunchEnabled: z.boolean().optional(),
  siteLogoUrl: z.string().optional(),
  title: z.string().optional().alias('t'),
  url: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
    error: e => `${e.input} is not a valid SharePoint Online site URL.`
  }).alias('u'),
  footerEnabled: z.boolean().optional(),
  navAudienceTargetingEnabled: z.boolean().optional(),
  searchScope: z.preprocess(value => String(value).toLowerCase(), z.enum(['defaultscope', 'tenant', 'hub', 'site'])).optional(),
  welcomePage: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoWebSetCommand extends SpoCommand {
  private static searchScopeOptions: string[] = ['defaultscope', 'tenant', 'hub', 'site'];

  public get name(): string {
    return commands.WEB_SET;
  }

  public get description(): string {
    return 'Updates subsite properties';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const payload: any = {};

    this.addUnknownOptionsToPayloadZod(payload, args.options);

    if (args.options.title) {
      payload.Title = args.options.title;
    }
    if (args.options.description) {
      payload.Description = args.options.description;
    }
    if (typeof args.options.siteLogoUrl !== 'undefined') {
      payload.SiteLogoUrl = args.options.siteLogoUrl;
    }
    if (typeof args.options.quickLaunchEnabled !== 'undefined') {
      payload.QuickLaunchEnabled = args.options.quickLaunchEnabled;
    }
    if (typeof args.options.headerEmphasis !== 'undefined') {
      payload.HeaderEmphasis = Number(args.options.headerEmphasis);
    }
    if (typeof args.options.headerLayout !== 'undefined') {
      payload.HeaderLayout = args.options.headerLayout === 'standard' ? 1 : 2;
    }
    if (typeof args.options.megaMenuEnabled !== 'undefined') {
      payload.MegaMenuEnabled = args.options.megaMenuEnabled;
    }
    if (typeof args.options.footerEnabled !== 'undefined') {
      payload.FooterEnabled = args.options.footerEnabled;
    }
    if (typeof args.options.navAudienceTargetingEnabled !== 'undefined') {
      payload.NavAudienceTargetingEnabled = args.options.navAudienceTargetingEnabled;
    }
    if (typeof args.options.searchScope !== 'undefined') {
      const searchScope = args.options.searchScope.toLowerCase();
      payload.SearchScope = SpoWebSetCommand.searchScopeOptions.indexOf(searchScope);
    }

    try {
      const requestOptions: CliRequestOptions = {
        url: `${args.options.url}/_api/web`,
        headers: {
          'content-type': 'application/json;odata=nometadata',
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json',
        data: payload
      };

      if (this.verbose) {
        await logger.logToStderr(`Updating properties of subsite ${args.options.url}...`);
      }

      await request.patch(requestOptions);

      if (typeof args.options.welcomePage !== 'undefined') {
        if (this.verbose) {
          await logger.logToStderr(`Setting welcome page to: ${args.options.welcomePage}...`);
        }

        const requestOptions: CliRequestOptions = {
          url: `${args.options.url}/_api/web/RootFolder`,
          headers: {
            'content-type': 'application/json;odata=nometadata',
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json',
          data: {
            WelcomePage: args.options.welcomePage
          }
        };

        await request.patch(requestOptions);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }
}

export default new SpoWebSetCommand();