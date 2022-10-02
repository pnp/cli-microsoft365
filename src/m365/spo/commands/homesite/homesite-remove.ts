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
  confirm?: boolean;
}

class SpoHomeSiteRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.HOMESITE_REMOVE;
  }

  public get description(): string {
    return 'Removes the current Home Site';
  }

  constructor() {
    super();
  
    this.#initTelemetry();
    this.#initOptions();
  }
  
  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        confirm: args.options.confirm || false
      });
    });
  }
  
  #initOptions(): void {
    this.options.unshift(
      {
        option: '--confirm'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {

    const removeHomeSite: () => Promise<void> = async (): Promise<void> => {
      try {
        const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
        const reqDigest = await spo.getRequestDigest(spoAdminUrl);

        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': reqDigest.FormDigestValue
          },
          data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="28" ObjectPathId="27" /><Method Name="RemoveSPHSite" Id="29" ObjectPathId="27" /></Actions><ObjectPaths><Constructor Id="27" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
        };

        const res = await request.post<string>(requestOptions);

        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          throw response.ErrorInfo.ErrorMessage;
        }
        else {
          logger.log(json[json.length - 1]);
        }
      }
      catch (err: any) {
        this.handleRejectedPromise(err);
      }
    };


    if (args.options.confirm) {
      await removeHomeSite();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the Home Site?`
      });

      if (result.continue) {
        await removeHomeSite();
      }
    }
  }
}

module.exports = new SpoHomeSiteRemoveCommand();