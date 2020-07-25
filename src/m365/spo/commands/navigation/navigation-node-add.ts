import request from '../../../../request';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

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
          cmd.log(chalk.green('DONE'));
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
        if (args.options.location !== 'QuickLaunch' &&
          args.options.location !== 'TopNavigationBar') {
          return `${args.options.location} is not a valid value for the location option. Allowed values are QuickLaunch|TopNavigationBar`;
        }
      }

      return true;
    };
  }
}

module.exports = new SpoNavigationNodeAddCommand();