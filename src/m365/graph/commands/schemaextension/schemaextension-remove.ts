import { Cli, Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  confirm?: boolean;
}

class GraphSchemaExtensionRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.SCHEMAEXTENSION_REMOVE;
  }

  public get description(): string {
    return 'Removes specified Microsoft Graph schema extension';
  }

  constructor() {
    super();
  
    this.#initTelemetry();
    this.#initOptions();
  }
  
  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        confirm: typeof args.options.confirm !== 'undefined'
      });
    });
  }
  
  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id <id>'
      },
      {
        option: '--confirm'
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const removeSchemaExtension: () => void = (): void => {
      if (this.verbose) {
        logger.logToStderr(`Removes specified Microsoft Graph schema extension with id '${args.options.id}'...`);
      }

      const requestOptions: any = {
        url: `${this.resource}/v1.0/schemaExtensions/${args.options.id}`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json'
        },
        responseType: 'json'
      };

      request.delete(requestOptions)
        .then(_ => cb(), (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));
    };
    if (args.options.confirm) {
      removeSchemaExtension();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the schema extension with ID ${args.options.id}?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeSchemaExtension();
        }
      });
    }
  }
}
module.exports = new GraphSchemaExtensionRemoveCommand();