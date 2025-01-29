import { Logger } from '../../../../cli/Logger.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import request, { CliRequestOptions } from '../../../../request.js';

class TenantReportSettingsGetCommand extends GraphCommand {
  public get name(): string {
    return commands.REPORT_SETTINGS_GET;
  }

  public get description(): string {
    return 'Get the tenant-level settings for Microsoft 365 reports';
  }

  public async commandAction(logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Getting tenant-level settings for Microsoft 365 reports...');
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/admin/reportSettings`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      const res = await request.get<any>(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TenantReportSettingsGetCommand();