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

interface Options extends GlobalOptions {
  id: string;
  title?: string;
  webTemplate?: string;
  siteScripts?: string;
  description?: string;
  previewImageUrl?: string;
  previewImageAltText?: string;
  version?: number | string;
  isDefault?: string;
}

class SpoSiteDesignSetCommand extends SpoCommand {
  public get name(): string {
    return `${commands.SITEDESIGN_SET}`;
  }

  public get description(): string {
    return 'Updates a site design with new values';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.title = typeof args.options.title !== 'undefined';
    telemetryProps.webTemplate = args.options.webTemplate;
    telemetryProps.siteScripts = typeof args.options.siteScripts !== 'undefined';
    telemetryProps.description = typeof args.options.description !== 'undefined';
    telemetryProps.previewImageUrl = typeof args.options.previewImageUrl !== 'undefined';
    telemetryProps.previewImageAltText = typeof args.options.previewImageAltText !== 'undefined';
    telemetryProps.version = typeof args.options.version !== 'undefined';
    telemetryProps.isDefault = typeof args.options.isDefault !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    this
      .getSpoUrl(cmd, this.debug)
      .then((spoUrl: string): Promise<any> => {
        const updateInfo: any = {
          Id: args.options.id
        };

        if (args.options.title) {
          updateInfo.Title = args.options.title;
        }
        if (args.options.description) {
          updateInfo.Description = args.options.description;
        }
        if (args.options.siteScripts) {
          updateInfo.SiteScriptIds = args.options.siteScripts.split(',').map(i => i.trim());
        }
        if (args.options.previewImageUrl) {
          updateInfo.PreviewImageUrl = args.options.previewImageUrl;
        }
        if (args.options.previewImageAltText) {
          updateInfo.PreviewImageAltText = args.options.previewImageAltText;
        }
        if (args.options.webTemplate) {
          updateInfo.WebTemplate = args.options.webTemplate === 'TeamSite' ? '64' : '68';
        }
        if (args.options.version) {
          updateInfo.Version = args.options.version;
        }
        if (typeof args.options.isDefault !== 'undefined') {
          updateInfo.IsDefault = args.options.isDefault === 'true';
        }

        const requestOptions: any = {
          url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign`,
          headers: {
            'content-type': 'application/json;charset=utf-8',
            accept: 'application/json;odata=nometadata',
            json: true
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
        description: 'The ID of the site design to update'
      },
      {
        option: '-t, --title [title]',
        description: 'The new display name of the updated site design'
      },
      {
        option: '-w, --webTemplate [webTemplate]',
        description: 'The new template to add the site design to. Allowed values TeamSite|CommunicationSite',
        autocomplete: ['TeamSite', 'CommunicationSite']
      },
      {
        option: '-s, --siteScripts [siteScripts]',
        description: 'Comma-separated list of new site script IDs. The scripts will run in the order listed'
      },
      {
        option: '-d, --description [description]',
        description: 'The new display description of updated site design'
      },
      {
        option: '-m, --previewImageUrl [previewImageUrl]',
        description: 'The new URL of a preview image. If none is specified SharePoint will use a generic image'
      },
      {
        option: '-a, --previewImageAltText [previewImageAltText]',
        description: 'The new alt text description of the image for accessibility'
      },
      {
        option: '-v, --version [version]',
        description: 'The new version number for the site design'
      },
      {
        option: '--isDefault [isDefault]',
        description: 'Set to true if the site design is applied as the default site design'
      },
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      if (args.options.webTemplate &&
        args.options.webTemplate !== 'TeamSite' &&
        args.options.webTemplate !== 'CommunicationSite') {
        return `${args.options.webTemplate} is not a valid web template type. Allowed values TeamSite|CommunicationSite`;
      }

      if (args.options.siteScripts) {
        const siteScripts = args.options.siteScripts.split(',');
        for (let i: number = 0; i < siteScripts.length; i++) {
          const trimmedId: string = siteScripts[i].trim();
          if (!Utils.isValidGuid(trimmedId)) {
            return `${trimmedId} is not a valid GUID`;
          }
        }
      }

      if (args.options.version &&
        typeof args.options.version !== 'number') {
        return `${args.options.version} is not a number`;
      }

      if (typeof args.options.isDefault !== 'undefined' &&
        args.options.isDefault !== 'true' &&
        args.options.isDefault !== 'false') {
        return `${args.options.isDefault} is not a valid boolean value`
      }

      return true;
    };
  }
}

module.exports = new SpoSiteDesignSetCommand();
