import * as chalk from 'chalk';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { SiteDesignRun } from './SiteDesignRun';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteDesignId?: string;
  webUrl: string;
}

class SpoSiteDesignRunListCommand extends SpoCommand {
  public get name(): string {
    return `${commands.SITEDESIGN_RUN_LIST}`;
  }

  public get description(): string {
    return 'Lists information about site designs applied to the specified site';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.siteDesignId = typeof args.options.siteDesignId !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const data: any = {};
    if (args.options.siteDesignId) {
      data.siteDesignId = args.options.siteDesignId;
    }

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRun`,
      headers: {
        accept: 'application/json;odata=nometadata',
        'content-type': 'application/json;odata=nometadata'
      },
      data: data,
      responseType: 'json'
    };

    request.post<{ value: SiteDesignRun[] }>(requestOptions)
      .then((res: { value: SiteDesignRun[] }): void => {
        if (args.options.output === 'json') {
          logger.log(res.value);
        }
        else {
          logger.log(res.value.map(d => {
            return {
              ID: d.ID,
              SiteDesignID: d.SiteDesignID,
              SiteDesignTitle: d.SiteDesignTitle,
              StartTime: new Date(parseInt(d.StartTime)).toLocaleString()
            };
          }));
        }

        if (this.verbose) {
          logger.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'The URL of the site for which to list applied site designs'
      },
      {
        option: '-i, --siteDesignId [siteDesignId]',
        description: 'The ID of the site design for which to display information'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
    if (isValidSharePointUrl !== true) {
      return isValidSharePointUrl;
    }

    if (args.options.siteDesignId) {
      if (!Utils.isValidGuid(args.options.siteDesignId)) {
        return `${args.options.siteDesignId} is not a valid GUID`;
      }
    }

    return true;
  }
}

module.exports = new SpoSiteDesignRunListCommand();