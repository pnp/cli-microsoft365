import { Group } from '@microsoft/microsoft-graph-types';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import aadCommands from '../../aadCommands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  displayName?: string;
  mailNickname?: string;
  force: boolean;
}

class EntraM365GroupRecycleBinItemRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.M365GROUP_RECYCLEBINITEM_REMOVE;
  }

  public get description(): string {
    return 'Permanently deletes a Microsoft 365 Group from the recycle bin in the current tenant';
  }

  public alias(): string[] | undefined {
    return [aadCommands.M365GROUP_RECYCLEBINITEM_REMOVE];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: typeof args.options.id !== 'undefined',
        displayName: typeof args.options.displayName !== 'undefined',
        mailNickname: typeof args.options.mailNickname !== 'undefined',
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
        option: '-d, --displayName [displayName]'
      },
      {
        option: '-m, --mailNickname [mailNickname]'
      },
      {
        option: '-f, --force'
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

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'displayName', 'mailNickname'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    await this.showDeprecationWarning(logger, aadCommands.M365GROUP_RECYCLEBINITEM_REMOVE, commands.M365GROUP_RECYCLEBINITEM_REMOVE);

    const removeGroup: () => Promise<void> = async (): Promise<void> => {
      try {
        const groupId = await this.getGroupId(args.options);

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/directory/deletedItems/${groupId}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeGroup();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the group '${args.options.id || args.options.displayName || args.options.mailNickname}'?` });

      if (result) {
        await removeGroup();
      }
    }
  }

  private async getGroupId(options: Options): Promise<string> {
    const { id, displayName, mailNickname } = options;

    if (id) {
      return id;
    }

    let filterValue: string = '';
    if (displayName) {
      filterValue = `displayName eq '${formatting.encodeQueryParameter(displayName)}'`;
    }

    if (mailNickname) {
      filterValue = `mailNickname eq '${formatting.encodeQueryParameter(mailNickname)}'`;
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/directory/deletedItems/Microsoft.Graph.Group?$filter=${filterValue}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: Group[] }>(requestOptions);
    const groups = response.value;

    if (groups.length === 0) {
      throw Error(`The specified group '${displayName || mailNickname}' does not exist.`);
    }

    if (groups.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', groups);
      const result = await cli.handleMultipleResultsFound<Group>(`Multiple groups with name '${displayName || mailNickname}' found.`, resultAsKeyValuePair);
      return result.id!;
    }

    return groups[0].id!;
  }
}

export default new EntraM365GroupRecycleBinItemRemoveCommand();