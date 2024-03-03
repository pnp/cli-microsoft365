import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  force?: boolean;
}

class PurviewRetentionLabelRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.RETENTIONLABEL_REMOVE;
  }

  public get description(): string {
    return 'Delete a retention label';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id <id>'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.id)) {
          return `'${args.options.id}' is not a valid GUID.`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removeRetentionLabel(args);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the retention label ${args.options.id}?` });

      if (result) {
        await this.removeRetentionLabel(args);
      }
    }
  }

  private async removeRetentionLabel(args: CommandArgs): Promise<void> {
    try {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/beta/security/labels/retentionLabels/${args.options.id}`,
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

export default new PurviewRetentionLabelRemoveCommand();