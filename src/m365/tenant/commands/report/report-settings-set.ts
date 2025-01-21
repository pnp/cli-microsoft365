import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { zod } from '../../../../utils/zod.js';

const options = globalOptionsZod
  .extend({
    displayConcealedNames: zod.alias('d', z.boolean())
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TenantReportSettingsSetCommand extends GraphCommand {
  public get name(): string {
    return commands.REPORT_SETTINGS_SET;
  }

  public get description(): string {
    return 'Update tenant-level settings for Microsoft 365 reports';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const { displayConcealedNames } = args.options;
      if (this.verbose) {
        await logger.logToStderr(`Updating report settings displayConcealedNames to ${displayConcealedNames}`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/admin/reportSettings`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json'
        },
        responseType: 'json',
        data: {
          displayConcealedNames: displayConcealedNames
        }
      };

      await request.patch(requestOptions);
    }
    catch (err) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TenantReportSettingsSetCommand();