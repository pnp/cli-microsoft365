import { z } from 'zod';
import { CommandError, globalOptionsZod } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import PowerAppsCommand from '../../../base/PowerAppsCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  name: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value is not a valid GUID.'
    })
    .alias('n'),
  force: z.boolean().optional().alias('f'),
  asAdmin: z.boolean().optional(),
  environmentName: z.string().optional().alias('e')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PaAppRemoveCommand extends PowerAppsCommand {
  public get name(): string {
    return commands.APP_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified Microsoft Power App';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => !opts.asAdmin || opts.environmentName, {
        message: 'When specifying the asAdmin option, the environment option is required as well.'
      })
      .refine(opts => !opts.environmentName || opts.asAdmin, {
        message: 'When specifying the environment option, the asAdmin option is required as well.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Removing Microsoft Power App ${args.options.name}...`);
    }

    const removePaApp = async (): Promise<void> => {
      let endpoint = `${this.resource}/providers/Microsoft.PowerApps`;
      if (args.options.asAdmin) {
        endpoint += `/scopes/admin/environments/${formatting.encodeQueryParameter(args.options.environmentName!)}`;
      }
      endpoint += `/apps/${formatting.encodeQueryParameter(args.options.name)}?api-version=2017-08-01`;

      const requestOptions: CliRequestOptions = {
        url: endpoint,
        fullResponse: true,
        headers: {
          accept: 'application/json'
        },
        responseType: 'json'
      };

      try {
        await request.delete(requestOptions);
      }
      catch (err: any) {
        if (err.response && err.response.status === 403) {
          throw new CommandError(`App '${args.options.name}' does not exist`);
        }
        else {
          this.handleRejectedODataJsonPromise(err);
        }
      }
    };

    if (args.options.force) {
      await removePaApp();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the Microsoft Power App ${args.options.name}?` });

      if (result) {
        await removePaApp();
      }
    }
  }
}

export default new PaAppRemoveCommand();