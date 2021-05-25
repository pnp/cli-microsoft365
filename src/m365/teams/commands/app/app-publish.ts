import * as fs from 'fs';
import * as path from 'path';
import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  filePath: string;
}

class TeamsAppPublishCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_PUBLISH;
  }

  public get description(): string {
    return 'Publishes Teams app to the organization\'s app catalog';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const fullPath: string = path.resolve(args.options.filePath);
    if (this.verbose) {
      logger.logToStderr(`Adding app '${fullPath}' to app catalog...`);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/appCatalogs/teamsApps`,
      headers: {
        "content-type": "application/zip",
        accept: 'application/json;odata.metadata=none'
      },
      data: fs.readFileSync(fullPath)
    };

    request
      .post<{ id: string; }>(requestOptions)
      .then((res: { id: string; }): void => {
        if (res && res.id) {
          logger.log(res.id);
        }

        cb();
      }, (res: any): void => this.handleRejectedODataJsonPromise(res, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-p, --filePath <filePath>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    const fullPath: string = path.resolve(args.options.filePath);

    if (!fs.existsSync(fullPath)) {
      return `File '${fullPath}' not found`;
    }

    if (fs.lstatSync(fullPath).isDirectory()) {
      return `Path '${fullPath}' points to a directory`;
    }

    return true;
  }
}

module.exports = new TeamsAppPublishCommand();