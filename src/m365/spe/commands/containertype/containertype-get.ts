import { Logger } from '../../../../cli/Logger.js';
import config from '../../../../config.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { ClientSvcResponse, ClientSvcResponseContents, FormDigestInfo, ContainerTypeProperties, spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  name?: string;
}

class SpeContainertypeGetCommand extends SpoCommand {
  public get name(): string {
    return commands.CONTAINERTYPE_GET;
  }

  public get description(): string {
    return 'Get a Container Type';
  }

  constructor() {
    super();
    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: typeof args.options.id !== 'undefined',
        name: typeof args.options.name !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --name [name]'
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      {
        options: ['id', 'name']
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id as string)) {
          return `${args.options.id} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('id', 'name');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);

      if (this.verbose) {
        await logger.logToStderr(`Getting the Container type...`);
      }

      const containerTypeId = await this.getContainerTypeId(args.options, spoAdminUrl, logger);
      const allContainerTypes = await this.getContainerTypeById(containerTypeId, spoAdminUrl, logger);
      await logger.log(allContainerTypes);
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }

  private async getContainerTypeById(containerTypeId: string, spoAdminUrl: string, logger: Logger): Promise<ContainerTypeProperties[]> {
    const formDigestInfo: FormDigestInfo = await spo.ensureFormDigest(spoAdminUrl, logger, undefined, this.debug);

    const requestOptions: CliRequestOptions = {
      url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': formDigestInfo.FormDigestValue
      },
      data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="49" ObjectPathId="48" /><Method Name="GetSPOContainerTypeById" Id="50" ObjectPathId="48"><Parameters><Parameter Type="Guid">{${containerTypeId}}</Parameter><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="48" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
    };

    const res: string = await request.post(requestOptions);
    const json: ClientSvcResponse = JSON.parse(res);
    const response: ClientSvcResponseContents = json[0];

    if (response.ErrorInfo) {
      throw response.ErrorInfo.ErrorMessage;
    }

    const containerTypes: ContainerTypeProperties[] = json[json.length - 1];
    return containerTypes;
  }

  private async getContainerTypeId(options: Options, spoAdminUrl: string, logger: Logger): Promise<string> {
    if (options.id) {
      return options.id;
    }

    const containerTypes: ContainerTypeProperties[] = await spo.getAllContainerTypes(spoAdminUrl, logger, this.debug);

    // Get id of the container type by name
    const containerType: ContainerTypeProperties | undefined = containerTypes.find(c => c.DisplayName === options.name);
    if (!containerType) {
      throw new Error(`Container type with name ${options.name} not found`);
    }

    const match = containerType.ContainerTypeId.match(/\/Guid\(([^)]+)\)\//);
    return match![1];
  }
}

export default new SpeContainertypeGetCommand();