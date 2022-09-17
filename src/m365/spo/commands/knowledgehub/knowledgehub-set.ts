import { Logger } from '../../../../cli';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { ClientSvcResponse, ClientSvcResponseContents, formatting, spo, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

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

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --url <url>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.url)
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
      const reqDigest = await spo.getRequestDigest(spoAdminUrl);

      if (this.verbose) {
        logger.logToStderr(`Adding ${args.options.url} as the Knowledge Hub Site`);
      }

      const requestOptions: any = {
        url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': reqDigest.FormDigestValue
        },
        data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions> <ObjectPath Id="35" ObjectPathId="34" /> <Method Name="SetKnowledgeHubSite" Id="36" ObjectPathId="34"> <Parameters> <Parameter Type="String">${formatting.escapeXml(args.options.url)}</Parameter> </Parameters> </Method> </Actions> <ObjectPaths> <Constructor Id="34" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /> </ObjectPaths></Request>`
      };

      const res = await request.post<string>(requestOptions);
      const json: ClientSvcResponse = JSON.parse(res);
      const response: ClientSvcResponseContents = json[0];
      if (response.ErrorInfo) {
        throw response.ErrorInfo.ErrorMessage;
      }
      
      logger.log(json[json.length - 1]);
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }
}

module.exports = new SpoKnowledgehubSetCommand();