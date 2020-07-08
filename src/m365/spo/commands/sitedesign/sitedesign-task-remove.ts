import request from '../../../../request';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import GlobalOptions from '../../../../GlobalOptions';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  taskId: string;
  confirm?: boolean;
}

class SpoSiteDesignTaskRemoveCommand extends SpoCommand {
  public get name(): string {
    return `${commands.SITEDESIGN_TASK_REMOVE}`;
  }

  public get description(): string {
    return 'Removes the specified site design scheduled for execution';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = args.options.confirm || false;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const removeSiteDesignTask: () => void = (): void => {
      this
        .getSpoUrl(cmd, this.debug)
        .then((spoUrl: string): Promise<any> => {
          const requestOptions: any = {
            url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.RemoveSiteDesignTask`,
            headers: {
              accept: 'application/json;odata=nometadata'
            },
            body: {
              taskId: args.options.taskId
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
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
    }
    if (args.options.confirm) {
      removeSiteDesignTask();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the site design task ${args.options.taskId}?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeSiteDesignTask();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --taskId <taskId>',
        description: 'The ID of the site design task to remove'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removing the site design task'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.taskId) {
        return 'Required parameter taskId missing';
      }

      if (!Utils.isValidGuid(args.options.taskId)) {
        return `${args.options.taskId} is not a valid GUID`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
    
    Removes the specified site design task with taskId
    ${chalk.grey('6ec3ca5b-d04b-4381-b169-61378556d76e')} scheduled for execution without
    prompting confirmation
      ${this.name} --taskId 6ec3ca5b-d04b-4381-b169-61378556d76e --confirm

    Removes the specified site design task with taskId
    ${chalk.grey('6ec3ca5b-d04b-4381-b169-61378556d76e')} scheduled for execution with
    prompt for confirmation before removing
      ${this.name} --taskId 6ec3ca5b-d04b-4381-b169-61378556d76e 

  More information:

    SharePoint site design and site script overview
      https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview
`);
  }
}

module.exports = new SpoSiteDesignTaskRemoveCommand();