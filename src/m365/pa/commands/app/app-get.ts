import { Cli, CommandOutput, Logger } from '../../../../cli';
import Command, {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import Utils from '../../../../Utils';
import request from '../../../../request';
import AzmgmtCommand from '../../../base/AzmgmtCommand';
import commands from '../../commands';
import * as paAppListCommand from '../app/app-list';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
}

class PaAppGetCommand extends AzmgmtCommand {
  public get name(): string {
    return commands.APP_GET;
  }

  public get description(): string {
    return 'Gets information about the specified Microsoft Power App';
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'displayName', 'description', 'appVersion', 'owner'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving information about Microsoft Power App ${args.options.name}...`);
    }

    let requestUrl: string = '';
    const isValidGuid: boolean = Utils.isValidGuid(args.options.name);
    if (isValidGuid) {
      requestUrl = `${this.resource}providers/Microsoft.PowerApps/apps/${encodeURIComponent(args.options.name)}?api-version=2016-11-01`;

      const requestOptions: any = {
        url: requestUrl,
        headers: {
          accept: 'application/json'
        },
        responseType: 'json'
      };

      request
        .get(requestOptions)
        .then((res: any): void => {
          logger.log(this.setProperties(res));
          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    }
    else {
      this
        .getApps(args, logger)
        .then((getAppsOutput: CommandOutput): void => {
          const allApps: any = JSON.parse(getAppsOutput.stdout);
          if (allApps.length > 0) {
            let app = allApps.find((a: any) => {
              return a.properties.displayName.toLowerCase() == `${args.options.name}`.toLowerCase();
            });
            if (!!app) {
              logger.log(this.setProperties(app));
            }
            else {
              if (this.verbose) {
                logger.logToStderr(`No app found with the name '${args.options.name}'`);
              }
            }
          }
          else {
            if (this.verbose) {
              logger.logToStderr('No apps found');
            }
          }
          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    }
  }

  private getApps(args: CommandArgs, logger: Logger): Promise<CommandOutput> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving all apps...`);
    }

    const options: GlobalOptions = {
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    return Cli.executeCommandWithOutput(paAppListCommand as Command, { options: { ...options, _: [] } });
  }

  private setProperties(app: any) {
    app.displayName = app.properties.displayName;
    app.description = app.properties.description || '';
    app.appVersion = app.properties.appVersion;
    app.owner = app.properties.owner.email || '';
    return app;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --name <name>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.name) {
      return 'Specify a valid name or GUID';
    }

    return true;
  }

}

module.exports = new PaAppGetCommand();
