import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteUrl: string;
  ids: string;
}

class SpoSiteRecycleBinItemRestoreCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_RECYCLEBINITEM_RESTORE;
  }

  public get description(): string {
    return 'Restores given items from the site recycle bin';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Restoring items from recycle bin at ${args.options.siteUrl}...`);
    }

    const requestUrl: string = `${args.options.siteUrl}/_api/site/RecycleBin/RestoreByIds`;

    const ids: string[] = Utils.splitAndTrim(args.options.ids);
    const idsChunks: string[][] = [];

    while (ids.length) {
      idsChunks.push(ids.splice(0, 20));
    }

    Promise.all(
      idsChunks.map((idsChunk: string[]) => {
        const requestOptions: any = {
          url: requestUrl,
          headers: {
            'accept': 'application/json;odata=nometadata',
            'content-type': 'application/json'
          },
          responseType: 'json',
          data: {
            ids: idsChunk
          }
        };

        return request.post(requestOptions);
      })
    ).then(_ => cb(), (rawRes: any): void => this.handleRejectedODataPromise(rawRes, logger, cb));

  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --siteUrl <siteUrl>'
      },
      {
        option: '-i, --ids <ids>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.siteUrl);
    if (isValidSharePointUrl !== true) {
      return isValidSharePointUrl;
    }

    const notValidGuidsList = Utils.splitAndTrim(args.options.ids).filter(id => !Utils.isValidGuid(id as string));

    if (notValidGuidsList.length > 0) {
      return `some GUIDs are not valid. Please double check: ${notValidGuidsList.map(guid => guid)}`;
    }

    return true;
  }
}

module.exports = new SpoSiteRecycleBinItemRestoreCommand();