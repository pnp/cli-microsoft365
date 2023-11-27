import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import config from '../../../../config.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { ClientSvcResponse, ClientSvcResponseContents, spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  force?: boolean;
}

class SpoKnowledgehubRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.KNOWLEDGEHUB_REMOVE;
  }

  public get description(): string {
    return 'Removes the Knowledge Hub Site setting for your tenant';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        force: (!(!args.options.force)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-f, --force'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeKnowledgehub = async (): Promise<void> => {
      try {
        const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
        const reqDigest = await spo.getRequestDigest(spoAdminUrl);

        if (this.verbose) {
          await logger.logToStderr(`Removing Knowledge Hub Site settings from your tenant`);
        }

        const requestOptions: CliRequestOptions = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': reqDigest.FormDigestValue
          },
          data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="29" ObjectPathId="28"/><Method Name="RemoveKnowledgeHubSite" Id="30" ObjectPathId="28"/></Actions><ObjectPaths><Constructor Id="28" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/></ObjectPaths></Request>`
        };

        const res = await request.post<string>(requestOptions);

        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          throw response.ErrorInfo.ErrorMessage;
        }

        await logger.log(json[json.length - 1]);
      }
      catch (err: any) {
        this.handleRejectedPromise(err);
      }
    };

    if (args.options.force) {
      if (this.debug) {
        await logger.logToStderr('Confirmation bypassed by entering confirm option. Removing Knowledge Hub Site setting...');
      }
      await removeKnowledgehub();
    }
    else {
      const result = await Cli.promptForConfirmation({ message: `Are you sure you want to remove Knowledge Hub Site from your tenant?` });

      if (result) {
        await removeKnowledgehub();
      }
    }
  }
}

export default new SpoKnowledgehubRemoveCommand();