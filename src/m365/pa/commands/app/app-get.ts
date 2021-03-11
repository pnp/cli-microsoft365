import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import Utils from '../../../../Utils';
import request from '../../../../request';
import AzmgmtCommand from '../../../base/AzmgmtCommand';
import commands from '../../commands';

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
      requestUrl = `${this.resource}providers/Microsoft.PowerApps/apps/${encodeURIComponent(args.options.name)}?api-version=2016-11-01`
    }
    else {
      requestUrl = `${this.resource}providers/Microsoft.PowerApps/apps?api-version=2016-11-01`
    }

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
        if (isValidGuid){
          res.displayName = res.properties.displayName;
          res.description = res.properties.description || '';
          res.appVersion = res.properties.appVersion;
          res.owner = res.properties.owner.email || '';

          logger.log(res);
        } else {
          if (res.value.length > 0) {
            let app = res.value.find((a: any)=> {
              return a.properties.displayName.toLowerCase() == args.options.name.toLowerCase();
            });
            if (!!app) {
              app.displayName = app.properties.displayName;
              app.description = app.properties.description || '';
              app.appVersion = app.properties.appVersion;
              app.owner = app.properties.owner.email || '';
              logger.log(app);
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
        }
        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
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
}

module.exports = new PaAppGetCommand();
