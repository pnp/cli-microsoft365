import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';

const options = globalOptionsZod
  .extend({
    asAdmin: z.boolean().optional()
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PpEnvironmentListCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.ENVIRONMENT_LIST;
  }

  public get description(): string {
    return 'Lists Microsoft Power Platform environments';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'displayName'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving list of Microsoft Power Platform environments...`);
    }

    let url: string = '';
    if (args.options.asAdmin) {
      url = `${this.resource}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments`;
    }
    else {
      url = `${this.resource}/providers/Microsoft.BusinessAppPlatform/environments`;
    }

    const requestOptions: CliRequestOptions = {
      url: `${url}?api-version=2020-10-01`,
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
      }

      await logger.log(res.value);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PpEnvironmentListCommand();
