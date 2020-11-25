import * as chalk from 'chalk';
import { Cli, Logger } from '../../../../cli';
import {
    CommandOption
} from '../../../../Command';
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
  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = typeof args.options.confirm !== 'undefined';
    return telemetryProps;
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
      .then((): void => {
        if (this.verbose) {
          logger.logToStderr(chalk.green('DONE'));
        }

        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));
  };
  if (args.options.confirm) {
    removeSchemaExtension();
  }
  else {
    Cli.prompt({
      type: 'confirm',
      name: 'continue',
      default: false,
      message: `Are you sure you want to remove the schema extension with ID ${args.options.id}?`,
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

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>',
        description: `The unique identifier for the schema extension definition`
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removing the specified schema extension'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}
module.exports = new GraphSchemaExtensionRemoveCommand();