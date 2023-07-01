import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

class SpoHideDefaultThemesGetCommand extends SpoCommand {
  public get name(): string {
    return commands.HIDEDEFAULTTHEMES_GET;
  }

  public get description(): string {
    return 'Gets the current value of the HideDefaultThemes setting';
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);

      if (this.verbose) {
        await logger.logToStderr(`Getting the current value of the HideDefaultThemes setting...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${spoAdminUrl}/_api/thememanager/GetHideDefaultThemes`,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const res = await request.post<any>(requestOptions);
      await logger.log(res.value);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoHideDefaultThemesGetCommand();