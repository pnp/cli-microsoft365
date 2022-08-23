import { Logger } from '../../../../cli';
import request from '../../../../request';
import AzmgmtCommand from '../../../base/AzmgmtCommand';
import commands from '../../commands';

class FlowEnvironmentListCommand extends AzmgmtCommand {
  public get name(): string {
    return commands.ENVIRONMENT_LIST;
  }

  public get description(): string {
    return 'Lists Microsoft Flow environments in the current tenant';
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'displayName'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving list of Microsoft Flow environments...`);
    }

    const requestOptions: any = {
      url: `${this.resource}providers/Microsoft.ProcessSimple/environments?api-version=2016-11-01`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    try {
      const res = await request.get<{ value: [{ name: string, displayName: string; properties: { displayName: string } }] }>(requestOptions);

      if (res.value && res.value.length > 0) {
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

module.exports = new FlowEnvironmentListCommand();