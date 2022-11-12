import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { formatting } from '../../../../utils/formatting';
import { odata } from '../../../../utils/odata';
import PowerAppsCommand from '../../../base/PowerAppsCommand';
import flowCommands from '../../../flow/commands';
import commands from '../../commands';
import { Connector } from './Connector';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environmentName: string;
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

  constructor() {
    super();

    this.#initOptions();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-e, --environmentName <environmentName>'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const url = `${this.resource}/providers/Microsoft.PowerApps/apis?api-version=2016-11-01&$filter=environment%20eq%20%27${formatting.encodeQueryParameter(args.options.environmentName)}%27%20and%20IsCustomApi%20eq%20%27True%27`;

    try {
      const connectors = await odata.getAllItems<Connector>(url);

      if (connectors.length > 0) {
        connectors.forEach(c => {
          c.displayName = c.properties.displayName;
        });

        logger.log(connectors);
      }
      else {
        if (this.verbose) {
          logger.logToStderr('No custom connectors found');
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new PaConnectorListCommand();