import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate,
  CommandError
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  type?: string;
  title?: string;
  alias?: string;
  description?: string;
  classification?: string;
  isPublic?: boolean;
  lcid?: number;
  url?: string;
  allowFileSharingForGuestUsers?: boolean;
  owners?: string;
  shareByEmailEnabled?: boolean;
  siteDesign?: string;
  siteDesignId?: string;
}

interface CreateGroupExResponse {
  DocumentsUrl: string;
  ErrorMessage: string;
  GroupId: string;
  SiteStatus: number;
  SiteUrl: string;
}

class SpoSiteAddCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_ADD;
  }

  public get description(): string {
    return 'Creates new modern site';
  }

  protected requiresTenantAdmin(): boolean {
    return false;
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    const isTeamSite: boolean = args.options.type === 'TeamSite';
    telemetryProps.siteType = args.options.type || 'TeamSite';
    telemetryProps.description = (!(!args.options.description)).toString();
    telemetryProps.classification = (!(!args.options.classification)).toString();
    telemetryProps.isPublic = args.options.isPublic || false;
    telemetryProps.lcid = args.options.lcid;
    telemetryProps.owners = typeof args.options.owners !== 'undefined';

    if (!isTeamSite) {
      telemetryProps.allowFileSharingForGuestUsers = args.options.allowFileSharingForGuestUsers || false;
      telemetryProps.shareByEmailEnabled = args.options.shareByEmailEnabled || false;
      telemetryProps.siteDesign = args.options.siteDesign;
      telemetryProps.siteDesignId = (!(!args.options.siteDesignId)).toString();
    }
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const isTeamSite: boolean = args.options.type !== 'CommunicationSite';

    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((accessToken: string): Promise<CreateGroupExResponse> => {
        if (args.options.allowFileSharingForGuestUsers && this.verbose) {
          cmd.log(vorpal.chalk.yellow(`Option 'allowFileSharingForGuestUsers' is deprecated. Please use 'shareByEmailEnabled' instead`));
        }

        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}`);
        }

        if (this.verbose) {
          cmd.log(`Creating new site...`);
        }

        let requestOptions: any = {}

        if (isTeamSite) {
          requestOptions = {
            url: `${auth.site.url}/_api/GroupSiteManager/CreateGroupEx`,
            headers: {
              authorization: `Bearer ${auth.service.accessToken}`,
              'content-type': 'application/json; odata=verbose; charset=utf-8',
              accept: 'application/json;odata=nometadata'
            },
            json: true,
            body: {
              displayName: args.options.title,
              alias: args.options.alias,
              isPublic: args.options.isPublic,
              optionalParams: {
                Description: args.options.description || '',
                CreationOptions: {
                  results: [],
                  Classification: args.options.classification || ''
                }
              }
            }
          };

          if (args.options.lcid) {
            requestOptions.body.optionalParams.CreationOptions.results.push(`SPSiteLanguage:${args.options.lcid}`);
          }

          if (args.options.owners) {
            requestOptions.body.optionalParams.Owners = {
              results: args.options.owners.split(',').map(o => o.trim())
            };
          }
        }
        else {
          let siteDesignId: string = '';
          if (args.options.siteDesignId) {
            siteDesignId = args.options.siteDesignId;
          }
          else {
            if (args.options.siteDesign) {
              switch (args.options.siteDesign) {
                case 'Topic':
                  siteDesignId = '00000000-0000-0000-0000-000000000000';
                  break;
                case 'Showcase':
                  siteDesignId = '6142d2a0-63a5-4ba0-aede-d9fefca2c767';
                  break;
                case 'Blank':
                  siteDesignId = 'f6cc5403-0d63-442e-96c0-285923709ffc';
                  break;
              }
            }
            else {
              siteDesignId = '00000000-0000-0000-0000-000000000000';
            }
          }

          requestOptions = {
            url: `${auth.site.url}/_api/SPSiteManager/Create`,
            headers: {
              authorization: `Bearer ${auth.service.accessToken}`,
              'content-type': 'application/json;odata=nometadata',
              accept: 'application/json;odata=nometadata'
            },
            json: true,
            body: {
              request: {
                Title: args.options.title,
                Url: args.options.url,
                ShareByEmailEnabled: args.options.shareByEmailEnabled || args.options.allowFileSharingForGuestUsers,
                Description: args.options.description || '',
                Classification: args.options.classification || '',
                WebTemplate: 'SITEPAGEPUBLISHING#0',
                SiteDesignId: siteDesignId
              }
            }
          };

          if (args.options.lcid) {
            requestOptions.body.request.Lcid = args.options.lcid;
          }
        }

        return request.post(requestOptions);
      })
      .then((res: CreateGroupExResponse): void => {
        if (isTeamSite) {
          if (res.ErrorMessage !== null) {
            cb(new CommandError(res.ErrorMessage));
            return;
          }
          else {
            cmd.log(res.SiteUrl);
          }
        }
        else {
          if (res.SiteStatus === 2) {
            cmd.log(res.SiteUrl);
          }
          else {
            cb(new CommandError('An error has occurred while creating the site'));
            return;
          }
        }
        cb();
      }, (res: any): void => this.handleRejectedODataJsonPromise(res, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--type [type]',
        description: 'Type of modern sites to list. Allowed values TeamSite|CommunicationSite, default TeamSite',
        autocomplete: ['TeamSite', 'CommunicationSite']
      },
      {
        option: '-u, --url <url>',
        description: 'Site URL (applies only to communication sites)'
      },
      {
        option: '-a, --alias <alias>',
        description: 'Site alias, used in the URL and in the team site group e-mail (applies only to team sites)'
      },
      {
        option: '-t, --title <title>',
        description: 'Site title'
      },
      {
        option: '-d, --description [description]',
        description: 'Site description'
      },
      {
        option: '-c, --classification [classification]',
        description: 'Site classification'
      },
      {
        option: '-l, --lcid [lcid]',
        description: 'Site language in the LCID format, eg. 1033 for en-US'
      },
      {
        option: '--isPublic',
        description: 'Determines if the associated group is public or not (applies only to team sites)'
      },
      {
        option: '--shareByEmailEnabled',
        description: 'Determines whether it\'s allowed to share file with guests (applies only to communication sites)'
      },
      {
        option: '--allowFileSharingForGuestUsers',
        description: `(deprecated. Use 'shareByEmailEnabled' instead) Determines whether it\'s allowed to share file with guests (applies only to communication sites)`
      },
      {
        option: '--siteDesign [siteDesign]',
        description: 'Type of communication site to create. Allowed values Topic|Showcase|Blank, default Topic. Specify either siteDesign or siteDesignId',
        autocomplete: ['Topic', 'Showcase', 'Blank']
      },
      {
        option: '--siteDesignId [siteDesignId]',
        description: 'Id of the custom site design to use to create the site. Specify either siteDesign or siteDesignId (applies only to communication sites)'
      },
      {
        option: '--owners [owners]',
        description: 'Comma-separated list of users to set as site owners (applies only to team sites)'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      let isTeamSite: boolean = true;

      if (args.options.type) {
        if (args.options.type !== 'TeamSite' &&
          args.options.type !== 'CommunicationSite') {
          return `${args.options.type} is not a valid modern site type. Allowed types are TeamSite and CommunicationSite`;
        }
        else {
          isTeamSite = args.options.type === 'TeamSite';
        }
      }

      if (!args.options.title) {
        return 'Required option title missing';
      }

      if (isTeamSite) {
        if (!args.options.alias) {
          return 'Required option alias missing';
        }
      }
      else {
        if (!args.options.url) {
          return 'Required option url missing';
        }

        const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.url);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.siteDesign) {
          if (args.options.siteDesign !== 'Topic' &&
            args.options.siteDesign !== 'Showcase' &&
            args.options.siteDesign !== 'Blank') {
            return `${args.options.siteDesign} is not a valid communication site type. Allowed types are Topic, Showcase and Blank`;
          }
        }

        if (args.options.siteDesignId) {
          if (!Utils.isValidGuid(args.options.siteDesignId)) {
            return `${args.options.siteDesignId} is not a valid GUID`;
          }
        }

        if (args.options.siteDesign && args.options.siteDesignId) {
          return 'Specify siteDesign or siteDesignId but not both';
        }

        if (args.options.owners) {
          return `You can only specify owners when creating Team Sites`;
        }
      }

      if (args.options.lcid) {
        if (isNaN(args.options.lcid)) {
          return `${args.options.lcid} is not a number`;
        }

        if (args.options.lcid < 0) {
          return `LCID must be greater than 0 (${args.options.lcid})`;
        }
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to a SharePoint Online site,
      using the ${chalk.blue(commands.LOGIN)} command.

  Remarks:
  
    To create a modern site, you have to first log in to SharePoint using the
    ${chalk.blue(commands.LOGIN)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.
   
  Examples:
  
    Create modern team site with private group
      ${chalk.grey(config.delimiter)} ${commands.SITE_ADD} --alias team1 --title Team 1

    Create modern team site with description and classification
      ${chalk.grey(config.delimiter)} ${commands.SITE_ADD} --type TeamSite --alias team1 --title Team 1 --description Site of team 1 --classification LBI

    Create modern team site with public group
      ${chalk.grey(config.delimiter)} ${commands.SITE_ADD} --type TeamSite --alias team1 --title Team 1 --isPublic

    Create modern team site using the Dutch language
      ${chalk.grey(config.delimiter)} ${commands.SITE_ADD} --alias team1 --title Team 1 --lcid 1043

    Create modern team site with the specified users as owners
      ${chalk.grey(config.delimiter)} ${commands.SITE_ADD} --alias team1 --title Team 1 --owners 'steve@contoso.com, bob@contoso.com'

    Create communication site using the Topic design
      ${chalk.grey(config.delimiter)} ${commands.SITE_ADD} --type CommunicationSite --url https://contoso.sharepoint.com/sites/marketing --title Marketing

    Create communication site using the Showcase design
      ${chalk.grey(config.delimiter)} ${commands.SITE_ADD} --type CommunicationSite --url https://contoso.sharepoint.com/sites/marketing --title Marketing --siteDesign Showcase

    Create communication site using a custom site design
      ${chalk.grey(config.delimiter)} ${commands.SITE_ADD} --type CommunicationSite --url https://contoso.sharepoint.com/sites/marketing --title Marketing --siteDesignId 99f410fe-dd79-4b9d-8531-f2270c9c621c

    Create communication site using the Blank design with description and classification
      ${chalk.grey(config.delimiter)} ${commands.SITE_ADD} --type CommunicationSite --url https://contoso.sharepoint.com/sites/marketing --title Marketing --description Site of the marketing department --classification MBI --siteDesign Blank
  
  More information
    
    Creating SharePoint Communication Site using REST
      https://docs.microsoft.com/en-us/sharepoint/dev/apis/communication-site-creation-rest
`);
  }
}

module.exports = new SpoSiteAddCommand();