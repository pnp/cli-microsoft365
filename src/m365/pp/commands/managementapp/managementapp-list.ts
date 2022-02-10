import { Logger } from '../../../../cli';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

class PpManagementAppListCommand extends GraphCommand {
  public get name(): string {
    return commands.MANAGEMENTAPP_LIST;
  }

  public get description(): string {
    return 'Lists management applications for Power Platform';
  }

  public commandAction(logger: Logger, args: any, cb: () => void): void {
    const requestOptions: any = {
      // This should be refactored once we implement a PowerPlatform base class as api.bap will differ between envs.
      url: `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/adminApplications?api-version=2020-06-01`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    request
      .get(requestOptions)
      .then((res: any): void => {
        if (res.value && res.value.length > 0) {
          logger.log(res.value);
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new PpManagementAppListCommand();
