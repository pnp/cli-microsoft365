import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import PowerAppsCommand from '../../../base/PowerAppsCommand.js';
import commands from '../../commands.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';

export const options = z.strictObject({ ...globalOptionsZod.shape });

class PaEnvironmentListCommand extends PowerAppsCommand {
  public get name(): string {
    return commands.ENVIRONMENT_LIST;
  }

  public get description(): string {
    return 'Lists Microsoft Power Apps environments in the current tenant';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'displayName'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving list of Microsoft Power Apps environments...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/providers/Microsoft.PowerApps/environments?api-version=2017-08-01`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    try {
      const res = await request.get<{ value: [{ name: string, displayName: string; properties: { displayName: string } }] }>(requestOptions);

      if (res.value.length > 0) {
        res.value.forEach(e => {
          e.displayName = e.properties.displayName;
        });
      }

      await logger.log(res.value);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PaEnvironmentListCommand();