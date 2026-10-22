import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { z } from 'zod';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  hideDefaultThemes: z.enum(['true', 'false']).transform(val => val === 'true')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoHideDefaultThemesSetCommand extends SpoCommand {
  public get name(): string {
    return commands.HIDEDEFAULTTHEMES_SET;
  }

  public get description(): string {
    return 'Sets the value of the HideDefaultThemes setting';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
      if (this.verbose) {
        await logger.logToStderr(`Setting the value of the HideDefaultThemes setting to ${args.options.hideDefaultThemes}...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${spoAdminUrl}/_api/thememanager/SetHideDefaultThemes`,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        data: {
          hideDefaultThemes: args.options.hideDefaultThemes
        },
        responseType: 'json'
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoHideDefaultThemesSetCommand();