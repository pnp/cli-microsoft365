import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  force?: boolean;
  id?: string;
  listId?: string;
  group?: string;
  listTitle?: string;
  title?: string;
  listUrl?: string;
  webUrl: string;
}

class SpoFieldRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.FIELD_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified list- or site column';
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
        listId: typeof args.options.listId !== 'undefined',
        listTitle: typeof args.options.listTitle !== 'undefined',
        listUrl: typeof args.options.listUrl !== 'undefined',
        id: typeof args.options.id !== 'undefined',
        group: typeof args.options.group !== 'undefined',
        title: typeof args.options.title !== 'undefined',
        force: (!(!args.options.force)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-l, --listTitle [listTitle]'
      },
      {
        option: '--listId [listId]'
      },
      {
        option: '--listUrl [listUrl]'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '-t, --title [title]'
      },
      {
        option: '-g, --group [group]'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
          return `${args.options.listId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'title', 'group'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let messageEnd: string;
    if (args.options.listId || args.options.listTitle) {
      messageEnd = `in list ${args.options.listId || args.options.listTitle}`;
    }
    else {
      messageEnd = `in site ${args.options.webUrl}`;
    }

    const removeField = async (listRestUrl: string, fieldId: string | undefined, title: string | undefined): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Removing field ${fieldId || title} ${messageEnd}...`);
      }

      let fieldRestUrl: string = '';
      if (fieldId) {
        fieldRestUrl = `/getbyid('${formatting.encodeQueryParameter(fieldId)}')`;
      }
      else {
        fieldRestUrl = `/getbyinternalnameortitle('${formatting.encodeQueryParameter(title as string)}')`;
      }

      const requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/web/${listRestUrl}fields${fieldRestUrl}`,
        method: 'POST',
        headers: {
          'X-HTTP-Method': 'DELETE',
          'If-Match': '*',
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      await request.post(requestOptions);
    };

    const prepareRemoval = async (): Promise<void> => {
      let listRestUrl: string = '';

      if (args.options.listId) {
        listRestUrl = `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/`;
      }
      else if (args.options.listTitle) {
        listRestUrl = `lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')/`;
      }
      else if (args.options.listUrl) {
        const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
        listRestUrl = `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/`;
      }

      if (args.options.group) {
        if (this.verbose) {
          await logger.logToStderr(`Retrieving fields assigned to group ${args.options.group}...`);
        }
        const requestOptions: CliRequestOptions = {
          url: `${args.options.webUrl}/_api/web/${listRestUrl}fields`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        try {
          const res = await request.get<any>(requestOptions);
          const filteredResults = res.value.filter((field: { Id: string | undefined, Group: string | undefined; }) => field.Group === args.options.group);
          if (this.verbose) {
            await logger.logToStderr(`${filteredResults.length} matches found...`);
          }

          const promises = [];
          for (let index = 0; index < filteredResults.length; index++) {
            promises.push(removeField(listRestUrl, filteredResults[index].Id, undefined));
          }

          await Promise.all(promises);
        }
        catch (err: any) {
          this.handleRejectedODataJsonPromise(err);
        }
      }
      else {
        try {
          await removeField(listRestUrl, args.options.id, args.options.title);
          // REST post call doesn't return anything
        }
        catch (err: any) {
          this.handleRejectedODataJsonPromise(err);
        }
      }
    };

    if (args.options.force) {
      await prepareRemoval();
    }
    else {
      const confirmMessage: string = `Are you sure you want to remove the ${args.options.group ? 'fields' : 'field'} ${args.options.id || args.options.title || 'from group ' + args.options.group} ${messageEnd}?`;

      const result = await Cli.promptForConfirmation(confirmMessage);

      if (result) {
        await prepareRemoval();
      }
    }
  }
}

export default new SpoFieldRemoveCommand();