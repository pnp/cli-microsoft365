import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

enum HeaderLayout {
  Standard = 1,
  Compact,
  Minimal,
  Extended
}

enum FooterLayout {
  Simple = 1,
  Extended
}

enum Alignment {
  Left = 0,
  Center,
  Right
}

enum Emphasis {
  Lightest = 0,
  Light,
  Dark,
  Darkest
}

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  siteUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  headerLayout: z.enum(['Standard', 'Compact', 'Minimal', 'Extended']).optional(),
  headerEmphasis: z.enum(['Lightest', 'Light', 'Dark', 'Darkest']).optional(),
  logoAlignment: z.enum(['Left', 'Center', 'Right']).optional(),
  footerLayout: z.enum(['Simple', 'Extended']).optional(),
  footerEmphasis: z.enum(['Lightest', 'Light', 'Dark', 'Darkest']).optional(),
  disableMegaMenu: z.boolean().optional(),
  hideTitleInHeader: z.boolean().optional(),
  disableFooter: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoSiteChromeSetCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_CHROME_SET;
  }

  public get description(): string {
    return 'Sets the chrome header and footer for the specified site';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const headerLayout = args.options.headerLayout ? HeaderLayout[args.options.headerLayout] : null;
    const headerEmphasis = args.options.headerEmphasis ? Emphasis[args.options.headerEmphasis] : null;
    const logoAlignment = args.options.logoAlignment ? Alignment[args.options.logoAlignment] : null;
    const footerLayout = args.options.footerLayout ? FooterLayout[args.options.footerLayout] : null;
    const footerEmphasis = args.options.footerEmphasis ? Emphasis[args.options.footerEmphasis] : null;
    const hideTitleInHeader = typeof args.options.hideTitleInHeader !== "undefined" ? args.options.hideTitleInHeader : null;
    const disableMegaMenu = typeof args.options.disableMegaMenu !== 'undefined' ? args.options.disableMegaMenu : null;
    const disableFooter = typeof args.options.disableFooter !== 'undefined' ? args.options.disableFooter : null;

    const body: any = {};
    if (headerLayout !== null) {
      body["headerLayout"] = headerLayout;
    }
    if (headerEmphasis !== null) {
      body["headerEmphasis"] = headerEmphasis;
    }
    if (logoAlignment !== null) {
      body["logoAlignment"] = logoAlignment;
    }
    if (footerLayout !== null) {
      body["footerLayout"] = footerLayout;
    }
    if (footerEmphasis !== null) {
      body["footerEmphasis"] = 3 - footerEmphasis; // Footer is inverted
    }
    if (hideTitleInHeader !== null) {
      body["hideTitleInHeader"] = hideTitleInHeader;
    }
    if (disableMegaMenu !== null) {
      body["megaMenuEnabled"] = !disableMegaMenu;
    }
    if (disableFooter !== null) {
      body["footerEnabled"] = !disableFooter;
    }

    const requestOptions: any = {
      url: `${args.options.siteUrl}/_api/web/SetChromeOptions`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      data: body,
      responseType: 'json'
    };

    try {
      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteChromeSetCommand();