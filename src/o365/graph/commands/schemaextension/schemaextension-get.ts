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
}

class GraphSchemaExtensionGet extends GraphCommand {
  public get name(): string {
    return `${commands.SCHEMAEXTENSION_GET}`;
  }

  public get description(): string {
    return 'Gets the properties of the specified schema extension definition';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
        if (this.verbose) {
          cmd.log(`Gets the properties of the specified schema extension definition with id '${args.options.id}'...`);
        }

        const requestOptions: any = {
          url: `${this.resource}/v1.0/schemaExtensions/${args.options.id}`,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'content-type': 'application/json'
          },
          json: true
        };

      request.get(requestOptions)
      .then((res: any): void => {

        cmd.log(res);

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>',
        description: `The unique identifier for the schema extension definition`
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

    To get properties of a schema extension definition, you have to pass the ID of the schema
    extension. 

  Examples:
  
    Gets properties of a schema extension definition with ID domain_myExtension
      ${chalk.grey(config.delimiter)} ${this.name} --id domain_myExtension`);
  }
}
module.exports = new GraphSchemaExtensionGet();