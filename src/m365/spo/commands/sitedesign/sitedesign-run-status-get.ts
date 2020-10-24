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

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  runId: string;
  webUrl: string;
}

class SpoSiteDesignRunStatusGetCommand extends SpoCommand {
  public get name(): string {
    return `${commands.SITEDESIGN_RUN_STATUS_GET}`;
  }

  public get description(): string {
    return 'Gets information about the site scripts executed for the specified site design';
  }

  public defaultProperties(): string[] | undefined {
    return ['ActionTitle', 'SiteScriptTitle', 'OutcomeText'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const data: any = {
      runId: args.options.runId
    };

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRunStatus`,
      headers: {
        accept: 'application/json;odata=nometadata',
        'content-type': 'application/json;odata=nometadata'
      },
      data: data,
      responseType: 'json'
    };

    request.post<{ value: any[] }>(requestOptions)
      .then((res: { value: any[] }): void => {
        logger.log(res.value);

        if (this.verbose) {
          logger.logToStderr(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'The URL of the site for which to get the information'
      },
      {
        option: '-i, --runId <runId>',
        description: 'ID of the site design applied to the site as retrieved using \'spo sitedesign run list\''
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

    if (!Utils.isValidGuid(args.options.runId)) {
      return `${args.options.runId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new SpoSiteDesignRunStatusGetCommand();