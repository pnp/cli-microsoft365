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

    const ids: string[] = this.splitIdsList(args.options.ids);
    const idsChunks: string[][] = [];

    while (ids.length) {
      idsChunks.push(ids.splice(0, 20));
    }

    idsChunks.forEach(async (idsChunk: string[], index: number) => {
      const requestOptions: any = {
        url: requestUrl,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json',
        data: {
          ids: idsChunk
        }
      };

      await request.post(requestOptions);

      if(index === idsChunks.length - 1) {
        cb();
      }
    });
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

    if (this.splitIdsList(args.options.ids).map(id => Utils.isValidGuid(id as string)).some(check => check !== true)) {
      return `some items in list ${args.options.ids} is not a valid GUID`;
    }

    return true;
  }

  private splitIdsList(ids: string): string[] {
    return ids.split(',').map(id => id.trim());
  }
}

module.exports = new SpoSiteRecycleBinItemRestoreCommand();