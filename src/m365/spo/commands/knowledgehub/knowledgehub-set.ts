import commands from '../../commands';
import config from '../../../../config';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import {
  CommandOption,
  CommandValidate,
  CommandError
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import Utils from '../../../../Utils';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
}

class SpoKnowledgehubSetCommand extends SpoCommand {
  public get name(): string {
    return commands.KNOWLEDGEHUB_SET;
  }

  public get description(): string {
    return 'Sets the Knowledge Hub Site for your tenant';
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
          cmd.log(`Adding ${args.options.url} as the Knowledge Hub Site`);
        }

        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': res.FormDigestValue
          },
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions> <ObjectPath Id="35" ObjectPathId="34" /> <Method Name="SetKnowledgeHubSite" Id="36" ObjectPathId="34"> <Parameters> <Parameter Type="String">${Utils.escapeXml(args.options.url)}</Parameter> </Parameters> </Method> </Actions> <ObjectPaths> <Constructor Id="34" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /> </ObjectPaths></Request>`
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
          cmd.log(json[json.length - 1]);

          if (this.verbose) {
            cmd.log(vorpal.chalk.green('DONE'));
          }

          cb();
        }
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));

  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>',
        description: 'URL of the site to set as Knowledge Hub'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.url) {
        return 'Required parameter url missing';
      }

      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.url);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      return true;
    };
  }

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.KNOWLEDGEHUB_SET).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} to use this command you have to have permissions to access
    the tenant admin site.
      
  Remarks:

    If the specified ${chalk.grey('url')} doesn't refer to an existing site collection,
    you will get a ${chalk.grey('404 - "404 FILE NOT FOUND"')} error.

  Examples:
  
    Sets the Knowledge Hub Site for your tenant
      ${this.name} --url https://contoso.sharepoint.com/sites/knowledgesite
`);
  }
}

module.exports = new SpoKnowledgehubSetCommand();