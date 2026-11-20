import fs from 'fs';
import path from 'path';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  filePath: z.string()
    .refine(val => {
      const fullPath = path.resolve(val);
      if (!fs.existsSync(fullPath)) {
        return false;
      }
      if (fs.lstatSync(fullPath).isDirectory()) {
        return false;
      }
      return true;
    }, {
      message: 'Specified file does not exist or points to a directory.'
    })
    .alias('p')
});

declare type Options = z.infer<typeof options>;
interface CommandArgs {
  options: Options;
}

class TeamsAppPublishCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_PUBLISH;
  }

  public get description(): string {
    return 'Publishes Teams app to the organization\'s app catalog';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const fullPath: string = path.resolve(args.options.filePath);
      if (this.verbose) {
        await logger.logToStderr(`Adding app '${fullPath}' to app catalog...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/appCatalogs/teamsApps`,
        headers: {
          'content-type': 'application/zip',
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: fs.readFileSync(fullPath)
      };

      const res = await request.post<any>(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TeamsAppPublishCommand();