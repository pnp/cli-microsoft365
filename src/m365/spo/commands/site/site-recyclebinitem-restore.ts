import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteUrl: string;
  id: string;
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

    const ids: string[] = args.options.id.split(',');
    const idsChunks = [];

    while (ids.length) {
      idsChunks.push(ids.splice(0, 20));
    }

    idsChunks.forEach((idsChunk) => {
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

      request
        .post(requestOptions)
        .then((): void => {
          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --siteUrl <siteUrl>'
      },
      {
        option: '-i, --id <id>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    return SpoCommand.isValidSharePointUrl(args.options.siteUrl);
  }
}

module.exports = new SpoSiteRecycleBinItemRestoreCommand();