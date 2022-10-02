import { Logger } from '../../../../cli/Logger';
import request from '../../../../request';
import PowerAppsCommand from '../../../base/PowerAppsCommand';
import commands from '../../commands';

class PaEnvironmentListCommand extends PowerAppsCommand {
  public get name(): string {
    return commands.ENVIRONMENT_LIST;
  }

  public get description(): string {
    return 'Lists Microsoft Power Apps environments in the current tenant';
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'displayName'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving list of Microsoft Power Apps environments...`);
    }

    const requestOptions: any = {
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

        logger.log(res.value);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new PaEnvironmentListCommand();