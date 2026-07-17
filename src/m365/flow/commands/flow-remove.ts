import { z } from 'zod';
import { cli } from '../../../cli/cli.js';
import { Logger } from '../../../cli/Logger.js';
import { globalOptionsZod } from '../../../Command.js';
import request, { CliRequestOptions } from '../../../request.js';
import { formatting } from '../../../utils/formatting.js';
import commands from '../commands.js';
import PowerAutomateCommand from '../../base/PowerAutomateCommand.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  name: z.uuid().alias('n'),
  environmentName: z.string().alias('e'),
  asAdmin: z.boolean().optional(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class FlowRemoveCommand extends PowerAutomateCommand {
  public get name(): string {
    return commands.REMOVE;
  }

  public get description(): string {
    return 'Removes the specified Microsoft Flow';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Removing Microsoft Flow ${args.options.name}...`);
    }

    const removeFlow = async (): Promise<void> => {
      const requestOptions: CliRequestOptions = {
        url: `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/${args.options.asAdmin ? 'scopes/admin/' : ''}environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.name)}?api-version=2016-11-01`,
        fullResponse: true,
        headers: {
          accept: 'application/json'
        },
        responseType: 'json'
      };

      try {
        const rawRes = await request.delete<any>(requestOptions);
        // handle 204 and throw error message to cmd when invalid flow id is passed
        // https://github.com/pnp/cli-microsoft365/issues/1063#issuecomment-537218957

        if (rawRes.statusCode === 204) {
          throw `Error: Resource '${args.options.name}' does not exist in environment '${args.options.environmentName}'`;
        }
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };
    if (args.options.force) {
      await removeFlow();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the Microsoft Flow ${args.options.name}?` });

      if (result) {
        await removeFlow();
      }
    }
  }
}

export default new FlowRemoveCommand();