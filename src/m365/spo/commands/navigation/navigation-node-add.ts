import request from '../../../../request';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
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
  isExternal?: boolean;
  location?: string;
  parentNodeId?: number;
  title: string;
  url: string;
  webUrl: string;
}

class SpoNavigationNodeAddCommand extends SpoCommand {
  public get name(): string {
    return `${commands.NAVIGATION_NODE_ADD}`;
  }

  public get description(): string {
    return 'Adds a navigation node to the specified site navigation';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.isExternal = args.options.isExternal;
    telemetryProps.parentNodeId = typeof args.options.parentNodeId !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log(`Adding navigation node...`);
    }

    const nodesCollection: string = args.options.parentNodeId ?
      `GetNodeById(${args.options.parentNodeId})/Children` :
      (args.options.location as string).toLowerCase();

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web/navigation/${nodesCollection}`,
      headers: {
        accept: 'application/json;odata=nometadata',
        'content-type': 'application/json;odata=nometadata'
      },
      body: {
        Title: args.options.title,
        Url: args.options.url,
        IsExternal: args.options.isExternal === true
      },
      json: true
    };

    request
      .post(requestOptions)
      .then((res: any): void => {
        cmd.log(res);

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'Absolute URL of the site to which navigation should be modified'
      },
      {
        option: '-l, --location <location>',
        description: 'Navigation type where the node should be added. Available options: QuickLaunch|TopNavigationBar',
        autocomplete: ['QuickLaunch', 'TopNavigationBar']
      },
      {
        option: '-t, --title <title>',
        description: 'Navigation node title'
      },
      {
        option: '--url <url>',
        description: 'Navigation node URL'
      },
      {
        option: '--parentNodeId [parentNodeId]',
        description: 'ID of the node below which the node should be added'
      },
      {
        option: '--isExternal',
        description: 'Set, if the navigation node points to an external URL'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.webUrl) {
        return 'Required option webUrl missing';
      }

      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      if (args.options.parentNodeId) {
        if (isNaN(args.options.parentNodeId)) {
          return `${args.options.parentNodeId} is not a number`;
        }
      }
      else {
        if (!args.options.location) {
          return 'Required option location missing';
        }
        else {
          if (args.options.location !== 'QuickLaunch' &&
            args.options.location !== 'TopNavigationBar') {
            return `${args.options.location} is not a valid value for the location option. Allowed values are QuickLaunch|TopNavigationBar`;
          }
        }
      }

      if (!args.options.title) {
        return 'Required option title missing';
      }

      if (!args.options.url) {
        return 'Required option url missing';
      }

      return true;
    };
  }

  public commandHelp(args: CommandArgs, log: (message: string) => void): void {
    log(vorpal.find(commands.NAVIGATION_NODE_ADD).helpInformation());
    log(
      `  Examples:
  
    Add a navigation node pointing to a SharePoint page to the top navigation
      ${commands.NAVIGATION_NODE_ADD} --webUrl https://contoso.sharepoint.com/sites/team-a --location TopNavigationBar --title About --url /sites/team-s/sitepages/about.aspx

    Add a navigation node pointing to an external page to the quick launch
      ${commands.NAVIGATION_NODE_ADD} --webUrl https://contoso.sharepoint.com/sites/team-a --location QuickLaunch --title "About us" --url https://contoso.com/about-us --isExternal

    Add a navigation node below an existing node
      ${commands.NAVIGATION_NODE_ADD} --webUrl https://contoso.sharepoint.com/sites/team-a --parentNodeId 2010 --title About --url /sites/team-s/sitepages/about.aspx
`);
  }
}

module.exports = new SpoNavigationNodeAddCommand();