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
  title: string;
  webTemplate: string;
  siteScripts: string;
  description?: string;
  previewImageUrl?: string;
  previewImageAltText?: string;
  isDefault?: boolean;
}

class SpoSiteDesignAddCommand extends SpoCommand {
  public get name(): string {
    return `${commands.SITEDESIGN_ADD}`;
  }

  public get description(): string {
    return 'Adds site design for creating modern sites';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.webTemplate = args.options.webTemplate;
    telemetryProps.numSiteScripts = args.options.siteScripts.split(',').length;
    telemetryProps.description = (!(!args.options.description)).toString();
    telemetryProps.previewImageUrl = (!(!args.options.previewImageUrl)).toString();
    telemetryProps.previewImageAltText = (!(!args.options.previewImageAltText)).toString();
    telemetryProps.isDefault = args.options.isDefault || false;
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
        const info: any = {
          Title: args.options.title,
          WebTemplate: args.options.webTemplate === 'TeamSite' ? '64' : '68',
          SiteScriptIds: args.options.siteScripts.split(',').map(i => i.trim())
        };

        if (args.options.description) {
          info.Description = args.options.description;
        }
        if (args.options.previewImageUrl) {
          info.PreviewImageUrl = args.options.previewImageUrl;
        }
        if (args.options.previewImageAltText) {
          info.PreviewImageAltText = args.options.previewImageAltText;
        }
        if (args.options.isDefault) {
          info.IsDefault = true;
        }

        const requestOptions: any = {
          url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign`,
          headers: {
            'X-RequestDigest': res.FormDigestValue,
            'content-type': 'application/json;charset=utf-8',
            accept: 'application/json;odata=nometadata'
          },
          body: { info: info },
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
        option: '-t, --title <title>',
        description: 'The display name of the site design'
      },
      {
        option: '-w, --webTemplate <webTemplate>',
        description: 'Identifies which base template to add the design to. Allowed values TeamSite|CommunicationSite',
        autocomplete: ['TeamSite', 'CommunicationSite']
      },
      {
        option: '-s, --siteScripts <siteScripts>',
        description: 'Comma-separated list of site script IDs. The scripts will run in the order listed'
      },
      {
        option: '-d, --description [description]',
        description: 'The display description of site design'
      },
      {
        option: '-m, --previewImageUrl [previewImageUrl]',
        description: 'The URL of a preview image. If none is specified SharePoint will use a generic image'
      },
      {
        option: '-a, --previewImageAltText [previewImageAltText]',
        description: 'The alt text description of the image for accessibility'
      },
      {
        option: '--isDefault',
        description: 'Set if the site design is applied as the default site design'
      },
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (args.options.webTemplate !== 'TeamSite' &&
        args.options.webTemplate !== 'CommunicationSite') {
        return `${args.options.webTemplate} is not a valid web template type. Allowed values TeamSite|CommunicationSite`;
      }

      const siteScripts = args.options.siteScripts.split(',');
      for (let i: number = 0; i < siteScripts.length; i++) {
        const trimmedId: string = siteScripts[i].trim();
        if (!Utils.isValidGuid(trimmedId)) {
          return `${trimmedId} is not a valid GUID`;
        }
      }

      return true;
    };
  }
}

module.exports = new SpoSiteDesignAddCommand();
