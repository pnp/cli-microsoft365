import { Logger } from '../../../../cli/Logger';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { ClientSvcResponse, ClientSvcResponseContents, ContextInfo, spo } from '../../../../utils/spo';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

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

  constructor() {
    super();

    this.#initOptions();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name <name>'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoAdminUrl: string = await  spo.getSpoAdminUrl(logger, this.debug);
      const res: ContextInfo = await spo.getRequestDigest(spoAdminUrl);
      if (this.verbose) {
        logger.logToStderr(`Getting ${args.options.name} theme from tenant...`);
      }

      const requestOptions: any = {
        url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': res.FormDigestValue
        },
        data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="12" ObjectPathId="11" /><ObjectPath Id="14" ObjectPathId="13" /><Query Id="15" ObjectPathId="13"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="11" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="13" ParentId="11" Name="GetTenantTheme"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.name)}</Parameter></Parameters></Method></ObjectPaths></Request>`
      };

      const processQuery: string = await request.post(requestOptions);

      const json: ClientSvcResponse = JSON.parse(processQuery);
      const contents: ClientSvcResponseContents = json.find(x => { return x['ErrorInfo']; });

      if (contents && contents.ErrorInfo) {
        throw contents.ErrorInfo.ErrorMessage || 'ClientSvc unknown error';
      }
      const json2: any = await Promise.resolve(json);
      const theme = json2[6];
      delete theme._ObjectType_;
      logger.log(theme);
    } 
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }
}

module.exports = new SpoThemeGetCommand();