import commands from '../../commands';
import config from '../../../../config';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import {
  CommandError
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions { }

class SpoKnowledgehubGetCommand extends SpoCommand {
  public get name(): string {
    return commands.KNOWLEDGEHUB_GET;
  }

  public get description(): string {
    return 'Gets the Knowledge Hub Site URL for your tenant';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    let spoAdminUrl: string = '';

    this
    .getSpoAdminUrl(cmd, this.debug)
    .then((_spoAdminUrl: string): Promise<ContextInfo> => {
      spoAdminUrl = _spoAdminUrl;
      return this.getRequestDigest(spoAdminUrl);
    })
    .then((res: ContextInfo): Promise<string> => {
      if (this.verbose) {
        cmd.log(`Getting the Knowledge Hub Site settings for your tenant`);
      }

      const requestOptions: any = {
        url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': res.FormDigestValue
        },
        body: `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}"><Actions><ObjectPath Id="5" ObjectPathId="4"/><Method Name="GetKnowledgeHubSite" Id="6" ObjectPathId="4"/></Actions><ObjectPaths><Constructor Id="4" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/></ObjectPaths></Request>`
      };

      return request.post(requestOptions);
    })
    .then((res: string): void => {
      const json: ClientSvcResponse = JSON.parse(res);
      const response: ClientSvcResponseContents = json[0];
      if (response.ErrorInfo) {
        cb(new CommandError(response.ErrorInfo.ErrorMessage));
      }
      else {
        const result: string = !json[json.length - 1] ? '' : json[json.length - 1];
        cmd.log(result);

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }
        cb();
      }
    }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.KNOWLEDGEHUB_GET).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} to use this command you have to have permissions to access
    the tenant admin site.

  Examples:
  
    Gets the Knowledge Hub Site URL for your tenant
      m365 ${this.name}
`);
  }
}

module.exports = new SpoKnowledgehubGetCommand();