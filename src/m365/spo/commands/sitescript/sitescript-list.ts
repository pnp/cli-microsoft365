import * as chalk from 'chalk';
import { Logger } from '../../../../cli';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ContextInfo } from '../../spo';

class SpoSiteScriptListCommand extends SpoCommand {
  public get name(): string {
    return `${commands.SITESCRIPT_LIST}`;
  }

  public get description(): string {
    return 'Lists site script available for use with site designs';
  }

  public commandAction(logger: Logger, args: {}, cb: () => void): void {
    let spoUrl: string = '';

    this
      .getSpoUrl(logger, this.debug)
      .then((_spoUrl: string): Promise<ContextInfo> => {
        spoUrl = _spoUrl;
        return this.getRequestDigest(spoUrl);
      })
      .then((res: ContextInfo): Promise<{ value: any[] }> => {
        const requestOptions: any = {
          url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScripts'`,
          headers: {
            'X-RequestDigest': res.FormDigestValue,
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        return request.post<{ value: any[] }>(requestOptions);
      })
      .then((res: { value: any[] }): void => {
        if (res.value && res.value.length > 0) {
          logger.log(res.value);
        }

        if (this.verbose) {
          logger.logToStderr(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new SpoSiteScriptListCommand();