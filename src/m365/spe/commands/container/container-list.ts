import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { SpeContainer, spe } from '../../../../utils/spe.js';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  containerTypeId?: string;
  containerTypeName?: string;
}

class SpeContainerListCommand extends GraphCommand {
  public get name(): string {
    return commands.CONTAINER_LIST;
  }

  public get description(): string {
    return 'Lists all Container Types';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'containerTypeId', 'createdDateTime'];
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
        containerTypeId: typeof args.options.containerTypeId !== 'undefined',
        containerTypeName: typeof args.options.containerTypeName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--containerTypeId [containerTypeId]'
      },
      {
        option: '--containerTypeName [containerTypeName]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.containerTypeId && !validation.isValidGuid(args.options.containerTypeId as string)) {
          return `${args.options.containerTypeId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['containerTypeId', 'containerTypeName'] });
  }

  #initTypes(): void {
    this.types.string.push('containerTypeId', 'containerTypeName');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving list of Containers...`);
      }

      const containerTypeId = await this.getContainerTypeId(logger, args.options);
      const allContainers = await odata.getAllItems<SpeContainer>(`${this.resource}/v1.0/storage/fileStorage/containers?$filter=containerTypeId eq ${formatting.encodeQueryParameter(containerTypeId)}`);
      await logger.log(allContainers);
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }

  private async getContainerTypeId(logger: Logger, options: Options): Promise<string> {
    if (options.containerTypeId) {
      return options.containerTypeId;
    }

    return spe.getContainerTypeIdByName(options.containerTypeName!);
  }
}

export default new SpeContainerListCommand();