import commands from '../../commands.js';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import { z } from 'zod';

const options = globalOptionsZod.strict();

class SpoBrandcenterSettingsListCommand extends SpoCommand {
  public get name(): string {
    return commands.BRANDCENTER_SETTINGS_LIST;
  }

  public get description(): string {
    return 'Lists the brand center configuration';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving brand center configuration...`);
    }

    try {
      const spoUrl: string = await spo.getSpoUrl(logger, this.debug);

      const requestOptions: CliRequestOptions = {
        url: `${spoUrl}/_api/Brandcenter/Configuration`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const res = await request.get(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

}

export default new SpoBrandcenterSettingsListCommand();