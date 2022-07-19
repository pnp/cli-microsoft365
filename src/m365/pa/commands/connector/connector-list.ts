import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils';
import PowerAppsCommand from '../../../base/PowerAppsCommand';
import flowCommands from '../../../flow/commands';
import commands from '../../commands';
import { Connector } from './Connector';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environment: string;
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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const url = `${this.resource}/providers/Microsoft.PowerApps/apis?api-version=2016-11-01&$filter=environment%20eq%20%27${encodeURIComponent(args.options.environment)}%27%20and%20IsCustomApi%20eq%20%27True%27`;

    odata
      .getAllItems<Connector>(url)
      .then((connectors: Connector[]): void => {
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

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-e, --environment <environment>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new PaConnectorListCommand();