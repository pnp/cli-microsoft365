import request from '../../../../request';
import commands from '../../commands';
import SpoCommand from '../../../base/SpoCommand';
import { ContextInfo } from '../../spo';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

class SpoSiteScriptListCommand extends SpoCommand {
  public get name(): string {
    return `${commands.SITESCRIPT_LIST}`;
  }

  public get description(): string {
    return 'Lists site script available for use with site designs';
  }

  public commandAction(cmd: CommandInstance, args: {}, cb: () => void): void {
    let spoUrl: string = '';

    this
      .getSpoUrl(cmd, this.debug)
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
          json: true
        };

        return request.post<{ value: any[] }>(requestOptions);
      })
      .then((res: { value: any[] }): void => {
        if (res.value && res.value.length > 0) {
          cmd.log(res.value);
        }

        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }
}

module.exports = new SpoSiteScriptListCommand();