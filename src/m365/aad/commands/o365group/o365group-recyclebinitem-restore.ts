import { Group } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { validation } from '../../../../utils/validation';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  displayName?: string;
  mailNickname?: string;
}

class AadO365GroupRecycleBinItemRestoreCommand extends GraphCommand {
  public get name(): string {
    return commands.O365GROUP_RECYCLEBINITEM_RESTORE;
  }

  public get description(): string {
    return 'Restores a deleted Microsoft 365 Group';
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
        mailNickname: typeof args.options.mailNickname !== 'undefined'
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
    if (this.verbose) {
      logger.logToStderr(`Restoring Microsoft 365 Group: ${args.options.id || args.options.displayName || args.options.mailNickname}...`);
    }

    try {
      const groupId = await this.getGroupId(args.options);
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/directory/deleteditems/${groupId}/restore`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json'
        },
        responseType: 'json'
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getGroupId(options: Options): Promise<string> {
    const { id, displayName, mailNickname } = options;

    if (id) {
      return Promise.resolve(id);
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

    return request
      .get<{ value: Group[] }>(requestOptions)
      .then((response: { value: Group[] }): Promise<string> => {
        const groups = response.value;

        if (groups.length === 0) {
          return Promise.reject(`The specified group '${displayName || mailNickname}' does not exist.`);
        }

        if (groups.length > 1) {
          return Promise.reject(`Multiple groups with name '${displayName || mailNickname}' found: ${groups.map(x => x.id).join(',')}.`);
        }

        return Promise.resolve(groups[0].id!);
      });
  }
}

module.exports = new AadO365GroupRecycleBinItemRestoreCommand();