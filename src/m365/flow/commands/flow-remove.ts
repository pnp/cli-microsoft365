import { Cli } from '../../../cli/Cli';
import { Logger } from '../../../cli/Logger';
import GlobalOptions from '../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../request';
import { formatting } from '../../../utils/formatting';
import { validation } from '../../../utils/validation';
import AzmgmtCommand from '../../base/AzmgmtCommand';
import commands from '../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environmentName: string;
  name: string;
  asAdmin?: boolean;
  force?: boolean;
}

class FlowRemoveCommand extends AzmgmtCommand {
  public get name(): string {
    return commands.REMOVE;
  }

  public get description(): string {
    return 'Removes the specified Microsoft Flow';
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
        asAdmin: !!args.options.asAdmin,
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name <name>'
      },
      {
        option: '-e, --environmentName <environmentName>'
      },
      {
        option: '--asAdmin'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.name)) {
          return `${args.options.name} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Removing Microsoft Flow ${args.options.name}...`);
    }

    const removeFlow = async (): Promise<void> => {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}providers/Microsoft.ProcessSimple/${args.options.asAdmin ? 'scopes/admin/' : ''}environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.name)}?api-version=2016-11-01`,
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
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the Microsoft Flow ${args.options.name}?`
      });

      if (result.continue) {
        await removeFlow();
      }
    }
  }
}

module.exports = new FlowRemoveCommand();