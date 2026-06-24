import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import PowerAppsCommand from '../../../base/PowerAppsCommand.js';
import flowCommands from '../../../flow/commands.js';
import commands from '../../commands.js';
import { Connector } from './Connector.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  environmentName: z.string().alias('e')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PaConnectorListCommand extends PowerAppsCommand {
  public get name(): string {
    return commands.CONNECTOR_LIST;
  }

  public get description(): string {
    return 'Lists custom connectors in the given environment';
  }

  public alias(): string[] | undefined {
    return [flowCommands.CONNECTOR_LIST];
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'displayName'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const url = `${this.resource}/providers/Microsoft.PowerApps/apis?api-version=2016-11-01&$filter=environment%20eq%20%27${formatting.encodeQueryParameter(args.options.environmentName)}%27%20and%20IsCustomApi%20eq%20%27True%27`;

    try {
      const connectors = await odata.getAllItems<Connector>(url);

      if (connectors.length > 0) {
        connectors.forEach(c => {
          c.displayName = c.properties.displayName;
        });
      }

      await logger.log(connectors);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PaConnectorListCommand();