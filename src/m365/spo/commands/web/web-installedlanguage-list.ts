import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { WebInstalledLanguagePropertiesCollection } from './WebPropertiesCollection.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
    error: e => `${e.input} is not a valid SharePoint Online site URL.`
  }).alias('u')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoWebInstalledLanguageListCommand extends SpoCommand {
  public get name(): string {
    return commands.WEB_INSTALLEDLANGUAGE_LIST;
  }

  public get description(): string {
    return 'Lists all installed languages on site';
  }

  public defaultProperties(): string[] | undefined {
    return ['DisplayName', 'LanguageTag', 'Lcid'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving all installed languages on site ${args.options.webUrl}...`);
    }

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web/RegionalSettings/InstalledLanguages`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const webInstalledLanguageProperties: WebInstalledLanguagePropertiesCollection = await request.get<WebInstalledLanguagePropertiesCollection>(requestOptions);
      await logger.log(webInstalledLanguageProperties.Items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoWebInstalledLanguageListCommand();