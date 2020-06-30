import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  hideDefaultThemes: string;
}

class SpoHideDefaultThemesSetCommand extends SpoCommand {
  public get name(): string {
    return commands.HIDEDEFAULTTHEMES_SET;
  }

  public get description(): string {
    return 'Sets the value of the HideDefaultThemes setting';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.hideDefaultThemes = args.options.hideDefaultThemes;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    this
      .getSpoAdminUrl(cmd, this.debug)
      .then((spoAdminUrl: string): Promise<void> => {
        if (this.verbose) {
          cmd.log(`Setting the value of the HideDefaultThemes setting to ${args.options.hideDefaultThemes}...`);
        }

        const requestOptions: any = {
          url: `${spoAdminUrl}/_api/thememanager/SetHideDefaultThemes`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          body: {
            hideDefaultThemes: args.options.hideDefaultThemes,
          },
          json: true
        };

        return request.post(requestOptions);
      })
      .then((): void => {
        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--hideDefaultThemes <hideDefaultThemes>',
        description: 'Set to true to hide default themes and to false to show them'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (typeof args.options.hideDefaultThemes === 'undefined') {
        return 'Required parameter hideDefaultThemes missing';
      }

      if (args.options.hideDefaultThemes !== 'false' &&
        args.options.hideDefaultThemes !== 'true') {
        return `${args.options.hideDefaultThemes} is not a valid boolean`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} to use this command you have to have permissions to access
    the tenant admin site.
  
  Examples:

    Hide default themes and allow users to use organization themes only
      ${commands.HIDEDEFAULTTHEMES_SET} --hideDefaultThemes true

  More information:

    SharePoint site theming
      https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-overview
      `);
  }
}

module.exports = new SpoHideDefaultThemesSetCommand();