import commands from '../../commands';
import request from '../../../../request';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption
} from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

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
          cmd.log(chalk.green('DONE'));
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
}
module.exports = new GraphSchemaExtensionGet();