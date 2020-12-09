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
import { SiteDesign } from './SiteDesign';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
}

class SpoSiteDesignGetCommand extends SpoCommand {
  private spoUrl: string = "";

  public get name(): string {
    return `${commands.SITEDESIGN_GET}`;
  }

  public get description(): string {
    return 'Gets information about the specified site design';
  }

  private getSiteDesignId(args: CommandArgs): Promise<string> {
    if (args.options.id) {
      return Promise.resolve(args.options.id);
    }

    const requestOptions: any = {
      url: `${this.spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesigns`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request
      .post<{ value: SiteDesign[] }>(requestOptions)
      .then(response => {
        const matchingSiteDesigns: SiteDesign[] = response.value.filter(x => x.Title === args.options.title);

        if (matchingSiteDesigns.length === 0) {
          return Promise.reject(`The specified site design does not exist`);
        }

        if (matchingSiteDesigns.length > 1) {
          return Promise.reject(`Multiple site designs with title ${args.options.title} found: ${matchingSiteDesigns.map(x => x.Id).join(', ')}`);
        }

        return Promise.resolve(matchingSiteDesigns[0].Id);
      });
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getSpoUrl(logger, this.debug)
      .then((_spoUrl: string): Promise<string> => {
        this.spoUrl = _spoUrl
        return this.getSiteDesignId(args);
      })
      .then((siteDesignId: string): Promise<string> => {
        const requestOptions: any = {
          url: `${this.spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignMetadata`,
          headers: {
            'content-type': 'application/json;charset=utf-8',
            accept: 'application/json;odata=nometadata'
          },
          data: { id: siteDesignId },
          responseType: 'json'
        };

        return request.post(requestOptions);
      })
      .then((res: any): void => {
        logger.log(res);

        if (this.verbose) {
          logger.logToStderr(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id [id]',
        description: 'Site design ID. Specify either id or title but not both'
      },
      {
        option: '--title [title]',
        description: 'Site design title. Specify either id or title but not both'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.id && args.options.title) {
      return 'Specify either id or title, but not both.';
    }

    if (!args.options.id && !args.options.title) {
      return 'Specify id or title, one is required';
    }

    if (args.options.id && !Utils.isValidGuid(args.options.id as string)) {
      return `${args.options.id} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new SpoSiteDesignGetCommand();