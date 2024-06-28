import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { cli } from '../../../../cli/cli.js';
import config from '../../../../config.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import { ClientSvcResponse, ClientSvcResponseContents, spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  force?: boolean;
}

class SpoTenantSiteArchiveCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_SITE_ARCHIVE;
  }

  public get description(): string {
    return 'Archives a site collection';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '-u, --url <url>' },
      { option: '-f, --force' }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.url)
    );
  }

  #initTypes(): void {
    this.types.string.push('url');
    this.types.boolean.push('force');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const archiveSite = async (): Promise<void> => {
      try {
        if (this.verbose) {
          await logger.logToStderr(`Archiving site ${args.options.url}...`);
        }

        const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.verbose);
        const reqDigest = await spo.getRequestDigest(spoAdminUrl);

        const requestOptions: CliRequestOptions = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': reqDigest.FormDigestValue
          },
          data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009">
                  <Actions>
                    <ObjectPath Id="2" ObjectPathId="1" />
                    <ObjectPath Id="4" ObjectPathId="3" />
                    <Query Id="5" ObjectPathId="3">
                      <Query SelectAllProperties="true">
                        <Properties />
                      </Query>
                    </Query>
                  </Actions>
                  <ObjectPaths>
                    <Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" />
                    <Method Id="3" ParentId="1" Name="ArchiveSiteByUrl">
                      <Parameters>
                        <Parameter Type="String">${args.options.url}</Parameter>
                      </Parameters>
                    </Method>
                  </ObjectPaths>
                  </Request>`
        };

        const res = await request.post<string>(requestOptions);

        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          throw response.ErrorInfo.ErrorMessage;
        }
      }
      catch (err: any) {
        this.handleRejectedPromise(err);
      }
    };

    if (args.options.force) {
      await archiveSite();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to archive site '${args.options.url}'?` });

      if (result) {
        await archiveSite();
      }
    }
  }
}

export default new SpoTenantSiteArchiveCommand();