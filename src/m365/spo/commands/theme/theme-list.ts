import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: GlobalOptions;
}

class SpoThemeListCommand extends SpoCommand {
  public get name(): string {
    return commands.THEME_LIST;
  }

  public get description(): string {
    return 'Retrieves the list of custom themes';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getSpoAdminUrl(logger, this.debug)
      .then((spoAdminUrl: string): Promise<any> => {
        if (this.verbose) {
          logger.log(`Retrieving themes from tenant store...`);
        }

        const requestOptions: any = {
          url: `${spoAdminUrl}/_api/thememanager/GetTenantThemingOptions`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        return request.post(requestOptions);
      })
      .then((rawRes: any): void => {
        const themePreviews: any[] = rawRes.themePreviews;
        if (themePreviews && themePreviews.length > 0) {
          if (args.options.output === 'json') {
            logger.log(themePreviews);
          }
          else {
            logger.log(themePreviews.map(a => {
              return { Name: a.name };
            }));
          }
        }
        else {
          if (this.verbose) {
            logger.log('No themes found');
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new SpoThemeListCommand();