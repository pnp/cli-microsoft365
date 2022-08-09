import { Cli, Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting, urlUtil, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  confirm?: boolean;
  fieldTitle?: string;
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
        confirm: (!(!args.options.confirm)).toString()
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
        option: '--fieldTitle [fieldTitle]'
      },
      {
        option: '-t, --title [title]'
      },
      {
        option: '-g, --group [group]'
      },
      {
        option: '--confirm'
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
    this.optionSets.push(['id', 'title', 'fieldTitle', 'group']);
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (args.options.fieldTitle) {
      args.options.title = args.options.fieldTitle;

      this.warn(logger, `Option 'fieldTitle' is deprecated. Please use 'title' instead.`);
    }

    let messageEnd: string;
    if (args.options.listId || args.options.listTitle) {
      messageEnd = `in list ${args.options.listId || args.options.listTitle}`;
    }
    else {
      messageEnd = `in site ${args.options.webUrl}`;
    }

    const removeField = (listRestUrl: string, fieldId: string | undefined, title: string | undefined): Promise<void> => {
      if (this.verbose) {
        logger.logToStderr(`Removing field ${fieldId || title} ${messageEnd}...`);
      }

      let fieldRestUrl: string = '';
      if (fieldId) {
        fieldRestUrl = `/getbyid('${formatting.encodeQueryParameter(fieldId)}')`;
      }
      else {
        fieldRestUrl = `/getbyinternalnameortitle('${formatting.encodeQueryParameter(title as string)}')`;
      }

      const requestOptions: any = {
        url: `${args.options.webUrl}/_api/web/${listRestUrl}fields${fieldRestUrl}`,
        method: 'POST',
        headers: {
          'X-HTTP-Method': 'DELETE',
          'If-Match': '*',
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      return request.post(requestOptions);
    };

    const prepareRemoval = (): void => {
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
          logger.logToStderr(`Retrieving fields assigned to group ${args.options.group}...`);
        }
        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/${listRestUrl}fields`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        request
          .get(requestOptions)
          .then((res: any): void => {
            const filteredResults = res.value.filter((field: { Id: string | undefined, Group: string | undefined; }) => field.Group === args.options.group);
            if (this.verbose) {
              logger.logToStderr(`${filteredResults.length} matches found...`);
            }

            const promises = [];
            for (let index = 0; index < filteredResults.length; index++) {
              promises.push(removeField(listRestUrl, filteredResults[index].Id, undefined));
            }

            Promise.all(promises).then(() => {
              cb();
            })
              .catch((err) => {
                this.handleRejectedODataJsonPromise(err, logger, cb);
              });
          }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
      }
      else {
        removeField(listRestUrl, args.options.id, args.options.title)
          .then((): void => {
            // REST post call doesn't return anything
            cb();
          }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
      }
    };

    if (args.options.confirm) {
      prepareRemoval();
    }
    else {
      const confirmMessage: string = `Are you sure you want to remove the ${args.options.group ? 'fields' : 'field'} ${args.options.id || args.options.title || 'from group ' + args.options.group} ${messageEnd}?`;

      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: confirmMessage
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          prepareRemoval();
        }
      });
    }
  }
}

module.exports = new SpoFieldRemoveCommand();