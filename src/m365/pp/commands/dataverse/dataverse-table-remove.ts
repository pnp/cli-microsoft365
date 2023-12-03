import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environmentName: string;
  name: string;
  force?: true;
  asAdmin?: boolean;
}

class PpDataverseTableRemoveCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.DATAVERSE_TABLE_REMOVE;
  }

  public get description(): string {
    return 'Removes a dataverse table in a given environment';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        asAdmin: !!args.options.asAdmin,
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-e, --environmentName <environmentName>'
      },
      {
        option: '-n, --name <name>'
      },
      {
        option: '--asAdmin'
      },
      {
        option: '-f, --force'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Removing a table for which the user is an admin...`);
    }

    if (args.options.force) {
      await this.removeDataverseTable(args.options);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the dataverse table ${args.options.name}?` });

      if (result) {
        await this.removeDataverseTable(args.options);
      }
    }
  }

  private async removeDataverseTable(options: Options): Promise<void> {
    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(options.environmentName, options.asAdmin);

      const requestOptions: CliRequestOptions = {
        url: `${dynamicsApiUrl}/api/data/v9.0/EntityDefinitions(LogicalName='${options.name}')`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      await request.delete(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PpDataverseTableRemoveCommand();