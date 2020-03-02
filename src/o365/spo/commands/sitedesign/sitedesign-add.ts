import request from '../../../../request';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { ContextInfo } from '../../spo';
import GlobalOptions from '../../../../GlobalOptions';

const vorpal: Vorpal = require('../../../../vorpal-init');

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
          cmd.log(vorpal.chalk.green('DONE'));
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
      if (!args.options.title) {
        return 'Required parameter title missing';
      }

      if (!args.options.webTemplate) {
        return 'Required parameter webTemplate missing';
      }

      if (args.options.webTemplate !== 'TeamSite' &&
        args.options.webTemplate !== 'CommunicationSite') {
        return `${args.options.webTemplate} is not a valid web template type. Allowed values TeamSite|CommunicationSite`;
      }

      if (!args.options.siteScripts) {
        return 'Required parameter siteScripts missing';
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

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    Each time you execute the ${chalk.blue(this.name)} command, it will create
    a new site design with a unique ID. Before creating a site design, be sure
    that another design with the same name doesn't already exist.

    When specifying IDs of site scripts to use with your site design, ensure
    that the IDs refer to existing site scripts or provisioning sites using
    the design will lead to unexpected results.

  Examples:

    Create new site design for provisioning modern team sites
      ${this.name} --title "Contoso team site" --webTemplate TeamSite --siteScripts "19b0e1b2-e3d1-473f-9394-f08c198ef43e,b2307a39-e878-458b-bc90-03bc578531d6"

    Create new default site design for provisioning modern communication sites
      ${this.name} --title "Contoso communication site" --webTemplate CommunicationSite --siteScripts "19b0e1b2-e3d1-473f-9394-f08c198ef43e" --isDefault

  More information:

    SharePoint site design and site script overview
      https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview

    Customize a default site design
      https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/customize-default-site-design

    Site design JSON schema
      https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-json-schema
`);
  }
}

module.exports = new SpoSiteDesignAddCommand();
