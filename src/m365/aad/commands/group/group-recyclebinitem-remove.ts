import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import request, { CliRequestOptions } from '../../../../request';
import { validation } from '../../../../utils/validation';
import { odata } from '../../../../utils/odata';
import { Group } from '@microsoft/microsoft-graph-types';
import { formatting } from '../../../../utils/formatting';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  displayName?: string;
  force?: boolean
}

class AadGroupRecycleBinItemRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.GROUP_RECYCLEBINITEM_REMOVE;
  }

  public get description(): string {
    return 'Removes a group from the tenant recycle bin';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
    this.#initTelemetry();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: args.options.id !== 'undefined',
        displayName: args.options.displayName !== 'undefined',
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --displayName [displayName]'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      {
        options: ['id', 'displayName']
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeGroupFromTenantRecycleBin = async (): Promise<void> => {
      if (this.verbose) {
        logger.logToStderr(`Removing group ${args.options.id || args.options.displayName} from the tenant recycle bin...`);
      }

      try {
        const requestUrl = `${this.resource}/v1.0/directory/deletedItems/${args.options.id || await this.getDeletedGroupIdByDisplayName(args.options.displayName!)}`;
        const requestOptions: CliRequestOptions = {
          url: requestUrl,
          headers: {
            'accept': 'application/json;odata.metadata=none'
          }
        };

        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeGroupFromTenantRecycleBin();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the group ${args.options.id || args.options.displayName} from the tenant recycle bin?`
      });

      if (result.continue) {
        await removeGroupFromTenantRecycleBin();
      }
    }
  }

  private async getDeletedGroupIdByDisplayName(displayName: string): Promise<string> {
    const groups = await odata.getAllItems<Group>(`${this.resource}/v1.0/directory/deletedItems/microsoft.graph.group?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'&$select=id`);

    if (!groups.length) {
      throw Error(`The specified deleted group '${displayName}' does not exist.`);
    }

    if (groups.length > 1) {
      throw Error(`Multiple deleted groups with name '${displayName}' found: ${groups.map(x => x.id).join(',')}.`);
    }

    return groups[0].id!;
  }
}

module.exports = new AadGroupRecycleBinItemRemoveCommand();