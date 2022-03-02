import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { spo } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { SiteDesign } from './SiteDesign';

interface CommandArgs {
  options: GlobalOptions;
}

class SpoSiteDesignListCommand extends SpoCommand {
  public get name(): string {
    return commands.SITEDESIGN_LIST;
  }

  public get description(): string {
    return 'Lists available site designs for creating modern sites';
  }

  public defaultProperties(): string[] | undefined {
    return ['Id', 'IsDefault', 'Title', 'Version', 'WebTemplate'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    spo
      .getSpoUrl(logger, this.debug)
      .then((spoUrl: string): Promise<{ value: SiteDesign[] }> => {
        const requestOptions: any = {
          url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesigns`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        return request.post(requestOptions);
      })
      .then((res: { value: SiteDesign[] }): void => {
        logger.log(res.value);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new SpoSiteDesignListCommand();