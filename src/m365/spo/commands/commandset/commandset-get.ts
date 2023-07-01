import { Logger } from '../../../../cli/Logger.js';
import os from 'os';
import GlobalOptions from '../../../../GlobalOptions.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  title?: string
  id?: string;
  clientSideComponentId?: string;
  scope?: string;
}

class SpoCommandSetGetCommand extends SpoCommand {
  private static readonly scopes: string[] = ['All', 'Site', 'Web'];
  private static readonly baseLocation: string = 'ClientSideExtension.ListViewCommandSet';
  private static readonly allowedCommandSetLocations: string[] = [SpoCommandSetGetCommand.baseLocation, `${SpoCommandSetGetCommand.baseLocation}.CommandBar`, `${SpoCommandSetGetCommand.baseLocation}.ContextMenu`];

  public get name(): string {
    return commands.COMMANDSET_GET;
  }

  public get description(): string {
    return 'Get a ListView Command Set that is added to a site.';
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
        title: typeof args.options.title !== 'undefined',
        id: typeof args.options.id !== 'undefined',
        clientSideComponentId: typeof args.options.clientSideComponentId !== 'undefined',
        scope: typeof args.options.scope !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-t, --title [title]'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '-c, --clientSideComponentId [clientSideComponentId]'
      },
      {
        option: '-s, --scope [scope]',
        autocomplete: SpoCommandSetGetCommand.scopes
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID.`;
        }

        if (args.options.clientSideComponentId && !validation.isValidGuid(args.options.clientSideComponentId)) {
          return `${args.options.clientSideComponentId} is not a valid GUID.`;
        }

        if (args.options.scope && SpoCommandSetGetCommand.scopes.indexOf(args.options.scope) < 0) {
          return `${args.options.scope} is not a valid scope. Valid scopes are ${SpoCommandSetGetCommand.scopes.join(', ')}`;
        }

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['title', 'id', 'clientSideComponentId'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Attempt to get a specific commandset by property ${args.options.title || args.options.id || args.options.clientSideComponentId}.`);
      }

      if (args.options.id) {
        const commandSet = await spo.getCustomActionById(args.options.webUrl, args.options.id, args.options.scope);

        if (commandSet === undefined) {
          throw `Command set with id ${args.options.id} can't be found.`;
        }
        else if (!SpoCommandSetGetCommand.allowedCommandSetLocations.some(allowedLocation => allowedLocation === commandSet.Location)) {
          throw `Custom action with id ${args.options.id} is not a command set.`;
        }
        await logger.log(commandSet);
      }
      else if (args.options.clientSideComponentId) {
        const filter = `${this.getBaseFilter()} ClientSideComponentId eq guid'${args.options.clientSideComponentId}'`;
        const commandSets = await spo.getCustomActions(args.options.webUrl, args.options.scope, filter);

        if (commandSets.length === 0) {
          throw `No command set with clientSideComponentId '${args.options.clientSideComponentId}' found.`;
        }
        await logger.log(commandSets[0]);
      }
      else if (args.options.title) {
        const filter = `${this.getBaseFilter()} Title eq '${formatting.encodeQueryParameter(args.options.title)}'`;
        const commandSets = await spo.getCustomActions(args.options.webUrl, args.options.scope, filter);

        if (commandSets.length === 1) {
          await logger.log(commandSets[0]);
        }
        else if (commandSets.length === 0) {
          throw `No command set with title '${args.options.title}' found.`;
        }
        else {
          throw `Multiple command sets with title '${args.options.title}' found. Please disambiguate using IDs: ${os.EOL}${commandSets.map(commandSet => `- ${commandSet.Id}`).join(os.EOL)}.`;
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getBaseFilter(): string {
    return `startswith(Location,'${SpoCommandSetGetCommand.baseLocation}') and`;
  }
}

export default new SpoCommandSetGetCommand();