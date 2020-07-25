import request from '../../../../request';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import GlobalOptions from '../../../../GlobalOptions';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  asTask: boolean;
  id: string;
  webUrl: string;
}

class SpoSiteDesignApplyCommand extends SpoCommand {
  public get name(): string {
    return `${commands.SITEDESIGN_APPLY}`;
  }

  public get description(): string {
    return 'Applies a site design to an existing site collection';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.asTask = args.options.asTask || false;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    this
      .getSpoUrl(cmd, this.debug)
      .then((spoUrl: string): Promise<any> => {
        const requestBody: any = {
          siteDesignId: args.options.id,
          webUrl: args.options.webUrl
        };

        const requestOptions: any = {
          url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.${args.options.asTask ? 'AddSiteDesignTask' : 'ApplySiteDesign'}`,
          headers: {
            'content-type': 'application/json;charset=utf-8',
            accept: 'application/json;odata=nometadata'
          },
          body: requestBody,
          json: true
        };

        return request.post(requestOptions);
      })
      .then((res: any): void => {
        if (res.value) {
          cmd.log(res.value);
        }
        else {
          cmd.log(res);
        }

        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>',
        description: 'The ID of the site design to apply'
      },
      {
        option: '-u, --webUrl <webUrl>',
        description: 'The URL of the site to apply the site design to'
      },
      {
        option: '--asTask',
        description: 'Apply site design as task. Required for large site designs'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      return SpoCommand.isValidSharePointUrl(args.options.webUrl);
    };
  }
}

module.exports = new SpoSiteDesignApplyCommand();