import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  asAdmin?: boolean;
}

class PpEnvironmentListCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.ENVIRONMENT_LIST;
  }

  public get description(): string {
    return 'Lists Microsoft Power Platform environments';
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'displayName'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving list of Microsoft Power Platform environments...`);
    }
    let url: string = '';
    if (args.options.asAdmin) {
      url = `${this.resource}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments`;
    }
    else {
      url = `${this.resource}/providers/Microsoft.BusinessAppPlatform/environments`;
    }

    const requestOptions: any = {
      url: `${url}?api-version=2020-10-01`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    request
      .get<{ value: [{ name: string, displayName: string; properties: { displayName: string } }] }>(requestOptions)
      .then((res: { value: [{ name: string, displayName: string; properties: { displayName: string } }] }): void => {
        if (res.value && res.value.length > 0) {
          res.value.forEach(e => {
            e.displayName = e.properties.displayName;
          });

          logger.log(res.value);
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-a, --asAdmin [teamId]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new PpEnvironmentListCommand();
