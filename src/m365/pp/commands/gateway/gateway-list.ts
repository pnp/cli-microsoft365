import { Logger } from '../../../../cli';
import request from '../../../../request';
import PowerBICommand from '../../../base/PowerBICommand';
import commands from '../../commands';

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

  public commandAction(logger: Logger, args: any, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving list of gateways for which the user is an admin...`);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/myorg/gateways`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    request
      .get<{ value: any[] }>(requestOptions)
      .then((res: { value: any[] }): void => {
        logger.log(res.value);
        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }
}

module.exports = new PpGatewayListCommand();
