import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import PowerBICommand from '../../../base/PowerBICommand.js';
import commands from '../../commands.js';

class PpGatewayListCommand extends PowerBICommand {
  public get name(): string {
    return commands.GATEWAY_LIST;
  }

  public get description(): string {
    return 'Returns a list of gateways for which the user is an admin';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'name'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving list of gateways for which the user is an admin...`);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/myorg/gateways`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      const res = await request.get<{ value: any[] }>(requestOptions);
      await logger.log(res.value);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PpGatewayListCommand();
