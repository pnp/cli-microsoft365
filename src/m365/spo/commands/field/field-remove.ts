import { Cli, Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { urlUtil, validation } from '../../../../utils';
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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = typeof args.options.listId !== 'undefined';
    telemetryProps.listTitle = typeof args.options.listTitle !== 'undefined';
    telemetryProps.listUrl = typeof args.options.listUrl !== 'undefined';
    telemetryProps.id = typeof args.options.id !== 'undefined';
    telemetryProps.group = typeof args.options.group !== 'undefined';
    telemetryProps.fieldTitle = typeof args.options.fieldTitle !== 'undefined';
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    let messageEnd: string;
    if (args.options.listId || args.options.listTitle) {
      messageEnd = `in list ${args.options.listId || args.options.listTitle}`;
    }
    else {
      messageEnd = `in site ${args.options.webUrl}`;
    }

    const removeField = (listRestUrl: string, fieldId: string | undefined, fieldTitle: string | undefined): Promise<void> => {
      if (this.verbose) {
        logger.logToStderr(`Removing field ${fieldId || fieldTitle} ${messageEnd}...`);
      }

      let fieldRestUrl: string = '';
      if (fieldId) {
        fieldRestUrl = `/getbyid('${encodeURIComponent(fieldId)}')`;
      }
      else {
        fieldRestUrl = `/getbyinternalnameortitle('${encodeURIComponent(fieldTitle as string)}')`;
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
        listRestUrl = `lists(guid'${encodeURIComponent(args.options.listId)}')/`;
      }
      else if (args.options.listTitle) {
        listRestUrl = `lists/getByTitle('${encodeURIComponent(args.options.listTitle as string)}')/`;
      }
      else if (args.options.listUrl) {
        const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
        listRestUrl = `GetList('${encodeURIComponent(listServerRelativeUrl)}')/`;
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
        removeField(listRestUrl, args.options.id, args.options.fieldTitle)
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
      const confirmMessage: string = `Are you sure you want to remove the ${args.options.group ? 'fields' : 'field'} ${args.options.id || args.options.fieldTitle || 'from group ' + args.options.group} ${messageEnd}?`;

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

  public options(): CommandOption[] {
    const options: CommandOption[] = [
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
        option: '-t, --fieldTitle [fieldTitle]'
      },
      {
        option: '-g, --group [group]'
      },
      {
        option: '--confirm'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
    if (isValidSharePointUrl !== true) {
      return isValidSharePointUrl;
    }

    if (!args.options.id && !args.options.fieldTitle && !args.options.group) {
      return 'Specify id, fieldTitle, or group. One is required';
    }

    if (args.options.id && !validation.isValidGuid(args.options.id)) {
      return `${args.options.id} is not a valid GUID`;
    }

    if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
      return `${args.options.listId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new SpoFieldRemoveCommand();