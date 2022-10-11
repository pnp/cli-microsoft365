import { Logger } from '../../../../cli/Logger';
import Command from '../../../../Command';
import { Cli, CommandOutput } from '../../../../cli/Cli';
import GlobalOptions from '../../../../GlobalOptions';
import { formatting } from '../../../../utils/formatting';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import { Options as SpoEventReceiverGetOptions } from './eventreceiver-get';
import commands from '../../commands';
import request from '../../../../request';
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
        scope: typeof args.options.scope !== 'undefined',
        id: typeof args.options.id !== 'undefined',
        name: typeof args.options.name !== 'undefined',
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
    this.optionSets.push(['name', 'id']);
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const getEventReceiversOutput = await this.getEventReceivers(args.options);
      const eventReceivers: EventReceiver[] = JSON.parse(getEventReceiversOutput.stdout);

      if (!eventReceivers.length) { 
        throw Error(`Specified event receiver with ${args.options.id !== undefined ? `id ${args.options.id}` : `name ${args.options.name}`} cannot be found`); 
      }

      if (eventReceivers.length > 1) { 
        throw Error(`Multiple eventreceivers with ${args.options.id !== undefined ? `id ${args.options.id} found` : `name ${args.options.name}, ids: ${eventReceivers.map(x => x.ReceiverId)} found`}`); 
      }

      if (args.options.confirm) {
        await this.removeEventReceiver(args.options);
      }
      else {
        const result = await Cli.prompt<{ continue: boolean }>({
          type: 'confirm',
          name: 'continue',
          default: false,
          message: `Are you sure you want to remove event receiver with ${args.options.id !== undefined ? `id ${args.options.id}` : `name ${args.options.name}`}?`
        });
  
        if (result.continue) {
          await this.removeEventReceiver(args.options);
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  public async removeEventReceiver(options: Options): Promise<void> {
    let requestUrl = `${options.webUrl}/_api/`;
    let listUrl: string = '';
    let filter: string = '?$filter=';

    if (options.listId) {
      listUrl = `lists(guid'${formatting.encodeQueryParameter(options.listId)}')/`;
    }
    else if (options.listTitle) {
      listUrl = `lists/getByTitle('${formatting.encodeQueryParameter(options.listTitle)}')/`;
    }
    else if (options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(options.webUrl, options.listUrl);
      listUrl = `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/`;
    }

    if (!options.scope || options.scope === 'web') {
      requestUrl += `web/${listUrl}eventreceivers`;
    }
    else {
      requestUrl += 'site/eventreceivers';
    }

    if (options.id) {
      filter += `receiverid eq (guid'${options.id}')`;
    }
    else {
      filter += `receivername eq '${options.name}'`;
    }
    const requestOptions: any = {
      url: requestUrl + filter,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    await request.delete<{ value: any[] }>(requestOptions);
  }

  private async getEventReceivers(options: Options): Promise<CommandOutput> {
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

    return await Cli.executeCommandWithOutput(getCommand as Command, { options: { ...getOptions, _: [] } });
  }
}

module.exports = new SpoEventreceiverRemoveCommand();