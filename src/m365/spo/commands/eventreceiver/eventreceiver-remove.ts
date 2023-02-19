import { Logger } from '../../../../cli/Logger';
import Command from '../../../../Command';
import { Cli } from '../../../../cli/Cli';
import GlobalOptions from '../../../../GlobalOptions';
import { formatting } from '../../../../utils/formatting';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import { Options as SpoEventReceiverGetOptions } from './eventreceiver-get';
import commands from '../../commands';
import request, { CliRequestOptions } from '../../../../request';
import { EventReceiver } from './EventReceiver';

const getCommand: Command = require('./eventreceiver-get');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
  listUrl?: string;
  scope?: string;
  id?: string;
  name?: string;
  confirm?: boolean;
}

class SpoEventreceiverRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.EVENTRECEIVER_REMOVE;
  }

  public get description(): string {
    return 'Removes event receivers for the specified web, site, or list.';
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
        scope: args.options.scope || 'web',
        id: typeof args.options.id !== 'undefined',
        name: typeof args.options.name !== 'undefined',
        confirm: !!args.options.confirm
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--listTitle [listTitle]'
      },
      {
        option: '--listId  [listId]'
      },
      {
        option: '--listUrl [listUrl]'
      },
      {
        option: '-n, --name [name]'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '-s, --scope [scope]',
        autocomplete: ['web', 'site']
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

        const listOptions: any[] = [args.options.listId, args.options.listTitle, args.options.listUrl];
        if (listOptions.some(item => item !== undefined) && listOptions.filter(item => item !== undefined).length > 1) {
          return `Specify either list id or title or list url`;
        }

        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
          return `${args.options.listId} is not a valid GUID`;
        }

        if (args.options.scope && ['web', 'site'].indexOf(args.options.scope) === -1) {
          return `${args.options.scope} is not a valid type value. Allowed values web|site.`;
        }

        if (args.options.scope && args.options.scope === 'site' && (args.options.listId || args.options.listUrl || args.options.listTitle)) {
          return 'Scope cannot be set to site when retrieving list event receivers.';
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['name', 'id'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.confirm) {
      await this.removeEventReceiver(args.options);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove event receiver with ${args.options.id ? `id ${args.options.id}` : `name ${args.options.name}`}?`
      });

      if (result.continue) {
        await this.removeEventReceiver(args.options);
      }
    }
  }

  public async removeEventReceiver(options: Options): Promise<void> {
    try {
      let requestUrl = `${options.webUrl}/_api/${options.scope || 'web'}`;

      if (options.listId) {
        requestUrl += `/lists('${options.listId}')`;
      }
      else if (options.listTitle) {
        requestUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(options.listTitle)}')`;
      }
      else if (options.listUrl) {
        const listServerRelativeUrl = urlUtil.getServerRelativePath(options.webUrl, options.listUrl);
        requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
      }

      const rerId = await this.getEventReceiverId(options);
      requestUrl += `/eventreceivers('${rerId}')`;

      const requestOptions: CliRequestOptions = {
        url: requestUrl,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      await request.delete(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getEventReceiverId(options: Options): Promise<string> {
    if (options.id) {
      return options.id;
    }

    const getOptions: SpoEventReceiverGetOptions = {
      webUrl: options.webUrl,
      listId: options.listId,
      listTitle: options.listTitle,
      listUrl: options.listUrl,
      scope: options.scope,
      id: options.id,
      name: options.name,
      debug: this.debug,
      verbose: this.verbose
    };

    const commandOutput = await Cli.executeCommandWithOutput(getCommand as Command, { options: { ...getOptions, _: [] } });
    const eventReceiver: EventReceiver = JSON.parse(commandOutput.stdout);

    return eventReceiver.ReceiverId;
  }
}

module.exports = new SpoEventreceiverRemoveCommand();