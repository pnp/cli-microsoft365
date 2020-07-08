import config from '../../../../config';
import commands from '../../commands';
import request from '../../../../request';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

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

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const removeSchemaExtension: () => void = (): void => {
        if (this.verbose) {
          cmd.log(`Removes specified Microsoft Graph schema extension with id '${args.options.id}'...`);
        }

        const requestOptions: any = {
          url: `${this.resource}/v1.0/schemaExtensions/${args.options.id}`,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'content-type': 'application/json'
          },
          json: true
        };

      request.delete(requestOptions)
      .then((): void => {
        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb));
  };
  if (args.options.confirm) {
    removeSchemaExtension();
  }
  else {
    cmd.prompt({
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

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.id) {
        return 'Required option id is missing';
      }
      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    To remove specified schema extension definition, you have to pass the ID of the schema
    extension. 

  Examples:
  
    Removes specified Microsoft Graph schema extension with ID domain_myExtension. Will prompt for confirmation
        ${chalk.grey(config.delimiter)} ${this.name} --id domain_myExtension
    
    Removes specified Microsoft Graph schema extension with ID domain_myExtension without prompt for confirmation
        ${chalk.grey(config.delimiter)} ${this.name} --id domain_myExtension --confirm
    `
    );    
  }
}
module.exports = new GraphSchemaExtensionRemoveCommand();