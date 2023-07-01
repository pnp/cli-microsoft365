import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

class SpoThemeListCommand extends SpoCommand {
  public get name(): string {
    return commands.THEME_LIST;
  }

  public get description(): string {
    return 'Retrieves the list of custom themes';
  }

  public defaultProperties(): string[] | undefined {
    return ['name'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const spoAdminUrl: string = await spo.getSpoAdminUrl(logger, this.debug);
      if (this.verbose) {
        await logger.logToStderr(`Retrieving themes from tenant store...`);
      }

      const requestOptions: any = {
        url: `${spoAdminUrl}/_api/thememanager/GetTenantThemingOptions`,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const rawRes: any = await request.post(requestOptions);
      const themePreviews: any[] = rawRes.themePreviews;
      if (themePreviews && themePreviews.length > 0) {
        await logger.log(themePreviews);
      }
      else {
        if (this.verbose) {
          await logger.logToStderr('No themes found');
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoThemeListCommand();