import request from '../../../../request';
import commands from '../../commands';
import SpoCommand from '../../../base/SpoCommand';
import { ContextInfo } from '../../spo';
import GlobalOptions from '../../../../GlobalOptions';
import { SiteDesign } from './SiteDesign';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: GlobalOptions;
}

class SpoSiteDesignListCommand extends SpoCommand {
  public get name(): string {
    return `${commands.SITEDESIGN_LIST}`;
  }

  public get description(): string {
    return 'Lists available site designs for creating modern sites';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    let spoUrl: string = '';

    this
      .getSpoUrl(cmd, this.debug)
      .then((_spoUrl: string): Promise<ContextInfo> => {
        spoUrl = _spoUrl;
        return this.getRequestDigest(spoUrl);
      })
      .then((res: ContextInfo): Promise<{ value: SiteDesign[] }> => {
        const requestOptions: any = {
          url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesigns`,
          headers: {
            'X-RequestDigest': res.FormDigestValue,
            accept: 'application/json;odata=nometadata'
          },
          json: true
        };

        return request.post(requestOptions);
      })
      .then((res: { value: SiteDesign[] }): void => {
        if (args.options.output === 'json') {
          cmd.log(res.value);
        }
        else {
          cmd.log(res.value.map(d => {
            return {
              Id: d.Id,
              IsDefault: d.IsDefault,
              Title: d.Title,
              Version: d.Version,
              WebTemplate: d.WebTemplate
            };
          }));
        }

        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }
}

module.exports = new SpoSiteDesignListCommand();