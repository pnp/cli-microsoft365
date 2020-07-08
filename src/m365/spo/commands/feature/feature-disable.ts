import SpoCommand from "../../../base/SpoCommand";
import GlobalOptions from "../../../../GlobalOptions";
import commands from "../../commands";
import {
  CommandTypes,
  CommandValidate,
  CommandOption
} from "../../../../Command";
import request from "../../../../request";

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  featureId: string;
  scope?: string;
  force: boolean;
}

class SpoFeatureDisableCommand extends SpoCommand {
  public get name(): string {
    return commands.FEATURE_DISABLE;
  }

  public get description(): string {
    return 'Disables feature for the specified site or web';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.scope = args.options.scope || 'web';
    telemetryProps.force = args.options.force || false;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    let scope: string | undefined = args.options.scope;
    let force: boolean = args.options.force;

    if (!scope) {
      scope = "web";
    }
    if (!force) {
      force = false;
    }

    if (this.verbose) {
      cmd.log(`Disabling feature '${args.options.featureId}' on scope '${scope}' for url '${args.options.url}' (force='${force}')...`);
    }

    const url: string = `${args.options.url}/_api/${scope}/features/remove(featureId=guid'${args.options.featureId}',force=${force})`;

    const requestOptions: any = {
      url: url,
      headers: {
        accept: 'application/json;odata=nometadata'
      }
    };

    request
      .post(requestOptions)
      .then((res: any): void => {
        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));

  }

  public types(): CommandTypes {
    return {
      string: ['scope', 's']
    };
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>',
        description: 'The URL of the site or web for which to disable the feature'
      },
      {
        option: '-f, --featureId <id>',
        description: 'The ID of the feature to disable'
      },
      {
        option: '-s, --scope [scope]',
        description: 'Scope of the Feature to disable. Allowed values `Site|Web`. Default `Web`',
        autocomplete: ['Site', 'Web']
      },
      {
        option: '--force',
        description: 'Specifies whether to continue if an error occurs when disabling the feature'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.url) {
        return 'Required parameter url missing';
      }

      if (!args.options.featureId) {
        return 'Required parameter featureId missing';
      }

      if (args.options.scope) {
        if (['site', 'web'].indexOf(args.options.scope.toLowerCase()) < 0) {
          return `${args.options.scope} is not a valid Feature scope. Allowed values are Site|Web`;
        }
      }

      return true;
    };
  }

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.FEATURE_DISABLE).helpInformation());
    log(
      `  Remarks:

    If the specified ${chalk.grey('url')} doesn't refer to an existing site collection,
    you will get a ${chalk.grey('"404 FILE NOT FOUND"')} error.
      
  Examples:
  
    Disable site feature
      ${this.name} --url https://contoso.sharepoint.com/sites/sales --featureId 915c240e-a6cc-49b8-8b2c-0bff8b553ed3 --scope Site

    Disable web feature (with force to ignore errors)
      ${this.name} --url https://contoso.sharepoint.com/sites/sales --featureId 00bfea71-5932-4f9c-ad71-1557e5751100 --scope Web --force
    `);
  }
}

module.exports = new SpoFeatureDisableCommand();