import * as chalk from 'chalk';
import { Logger } from '../../../../cli';
import {
    CommandOption
} from '../../../../Command';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ClientSvcResponse, ClientSvcResponseContents, ContextInfo } from '../../spo';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
}

class SpoThemeGetCommand extends SpoCommand {
  public get name(): string {
    return commands.THEME_GET;
  }

  public get description(): string {
    return 'Gets custom theme information';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    let spoAdminUrl: string = '';

    this
      .getSpoAdminUrl(logger, this.debug)
      .then((_spoAdminUrl: string): Promise<ContextInfo> => {
        spoAdminUrl = _spoAdminUrl;
        return this.getRequestDigest(spoAdminUrl);
      })
      .then((res: ContextInfo): Promise<string> => {
        if (this.verbose) {
          logger.logToStderr(`Getting ${args.options.name} theme from tenant...`);
        }

        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': res.FormDigestValue
          },
          data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="12" ObjectPathId="11" /><ObjectPath Id="14" ObjectPathId="13" /><Query Id="15" ObjectPathId="13"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="11" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="13" ParentId="11" Name="GetTenantTheme"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.name)}</Parameter></Parameters></Method></ObjectPaths></Request>`
        };

        return request.post(requestOptions);
      })
      .then((res: string): Promise<any> => {
        const json: ClientSvcResponse = JSON.parse(res);
        const contents: ClientSvcResponseContents = json.find(x => { return x['ErrorInfo']; });

        if (contents && contents.ErrorInfo) {
          return Promise.reject(contents.ErrorInfo.ErrorMessage || 'ClientSvc unknown error');
        }
        return Promise.resolve(json);
      })
      .then((json: any): void => {
        const theme = json[6];
        delete theme._ObjectType_;
        logger.log(theme);

        if (this.verbose) {
          logger.logToStderr(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [{
      option: '-n, --name <name>',
      description: 'The name of the theme to retrieve'
    }];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new SpoThemeGetCommand();