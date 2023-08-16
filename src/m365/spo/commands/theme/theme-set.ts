import { Logger } from '../../../../cli/Logger.js';
import config from '../../../../config.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { ClientSvcResponse, ClientSvcResponseContents, ContextInfo, spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  theme: string;
  isInverted: boolean;
}

class SpoThemeSetCommand extends SpoCommand {
  public get name(): string {
    return commands.THEME_SET;
  }

  public get description(): string {
    return 'Add or update a theme';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        inverted: (!(!args.options.isInverted)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name <name>'
      },
      {
        option: '-t, --theme <theme>'
      },
      {
        option: '--isInverted'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidTheme(args.options.theme)) {
          return 'The specified theme is not valid';
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoAdminUrl: string = await spo.getSpoAdminUrl(logger, this.debug);
      const res: ContextInfo = await spo.getRequestDigest(spoAdminUrl);
      const palette: any = JSON.parse(args.options.theme);

      if (this.debug) {
        await logger.logToStderr('');
        await logger.logToStderr('Palette');
        await logger.logToStderr(JSON.stringify(palette));
      }

      const isInverted: boolean = args.options.isInverted ? true : false;

      const requestOptions: any = {
        url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': res.FormDigestValue
        },
        data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><Method Name="UpdateTenantTheme" Id="11" ObjectPathId="9"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.name)}</Parameter><Parameter Type="String">{"isInverted":${isInverted},"name":"${formatting.escapeXml(args.options.name)}","palette":${JSON.stringify(palette)}}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="9" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/></ObjectPaths></Request>`
      };

      const processQuery: string = await request.post(requestOptions);
      const json: ClientSvcResponse = JSON.parse(processQuery);
      const contents: ClientSvcResponseContents = json.find(x => { return x['ErrorInfo']; });

      if (contents && contents.ErrorInfo) {
        throw contents.ErrorInfo.ErrorMessage || 'ClientSvc unknown error';
      }
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }
}

export default new SpoThemeSetCommand();