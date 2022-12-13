import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { ClientSvcResponse, ClientSvcResponseContents, spo } from '../../../../utils/spo';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  enabled: boolean;
  confirm?: boolean;
}

class SpoServicePrincipalSetCommand extends SpoCommand {
  public get name(): string {
    return commands.SERVICEPRINCIPAL_SET;
  }

  public get description(): string {
    return 'Enable or disable the service principal';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        enabled: args.options.enabled
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-e, --enabled <enabled>',
        autocomplete: ['true', 'false']
      },
      {
        option: '--confirm'
      }
    );
  }

  #initTypes(): void {
    this.types.boolean.push('enabled');
  }

  public alias(): string[] | undefined {
    return [commands.SP_SET];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {

    const toggleServicePrincipal: () => Promise<void> = async (): Promise<void> => {
      try {
        const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
        const reqDigest = await spo.getRequestDigest(spoAdminUrl);

        if (this.verbose) {
          logger.logToStderr(`${(args.options.enabled ? 'Enabling' : 'Disabling')} service principal...`);
        }

        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': reqDigest.FormDigestValue
          },
          data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="28" ObjectPathId="27" /><SetProperty Id="29" ObjectPathId="27" Name="AccountEnabled"><Parameter Type="Boolean">${args.options.enabled}</Parameter></SetProperty><Method Name="Update" Id="30" ObjectPathId="27" /><Query Id="31" ObjectPathId="27"><Query SelectAllProperties="true"><Properties><Property Name="AccountEnabled" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="27" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /></ObjectPaths></Request>`
        };

        const res = await request.post<string>(requestOptions);
        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          throw response.ErrorInfo.ErrorMessage;
        }
        else {
          const output: any = json[json.length - 1];
          delete output._ObjectType_;

          logger.log(output);
        }
      }
      catch (err: any) {
        this.handleRejectedPromise(err);
      }
    };

    if (args.options.confirm) {
      await toggleServicePrincipal();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to ${args.options.enabled ? 'enable' : 'disable'} the service principal?`
      });

      if (result.continue) {
        await toggleServicePrincipal();
      }
    }
  }
}

module.exports = new SpoServicePrincipalSetCommand();