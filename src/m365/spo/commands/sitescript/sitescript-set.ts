import request from '../../../../request';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { ContextInfo } from '../../spo';
import GlobalOptions from '../../../../GlobalOptions';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  title?: string;
  description?: string;
  version?: string;
  content?: string;
}

class SpoSiteScriptSetCommand extends SpoCommand {
  public get name(): string {
    return `${commands.SITESCRIPT_SET}`;
  }

  public get description(): string {
    return 'Updates existing site script';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.title = (!(!args.options.title)).toString();
    telemetryProps.description = (!(!args.options.description)).toString();
    telemetryProps.version = (!(!args.options.version)).toString();
    telemetryProps.content = (!(!args.options.content)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    let spoUrl: string = '';

    this
      .getSpoUrl(cmd, this.debug)
      .then((_spoUrl: string): Promise<ContextInfo> => {
        spoUrl = _spoUrl;
        return this.getRequestDigest(spoUrl);
      })
      .then((res: ContextInfo): Promise<string> => {
        const updateInfo: any = {
          Id: args.options.id
        };
        if (args.options.title) {
          updateInfo.Title = args.options.title;
        }
        if (args.options.description) {
          updateInfo.Description = args.options.description;
        }
        if (args.options.version) {
          updateInfo.Version = parseInt(args.options.version);
        }
        if (args.options.content) {
          updateInfo.Content = args.options.content;
        }

        const requestOptions: any = {
          url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteScript`,
          headers: {
            'X-RequestDigest': res.FormDigestValue,
            'content-type': 'application/json;charset=utf-8',
            accept: 'application/json;odata=nometadata'
          },
          body: { updateInfo: updateInfo },
          json: true
        };

        return request.post(requestOptions);
      })
      .then((res: any): void => {
        cmd.log(res);

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
        description: 'Site script ID'
      },
      {
        option: '-t, --title [title]',
        description: 'Site script title'
      },
      {
        option: '-d, --description [description]',
        description: 'Site script description'
      },
      {
        option: '-v, --version [version]',
        description: 'Site script version'
      },
      {
        option: '-c, --content [content]',
        description: 'JSON string containing the site script'
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

      if (args.options.version) {
        const version: number = parseInt(args.options.version);
        if (isNaN(version)) {
          return `${args.options.version} is not a number`;
        }
      }

      if (args.options.content) {
        try {
          JSON.parse(args.options.content);
        }
        catch (e) {
          return `Specified content value is not a valid JSON string. Error: ${e}`;
        }
      }

      return true;
    };
  }
}

module.exports = new SpoSiteScriptSetCommand();