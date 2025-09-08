import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { GroupExtended } from './GroupExtended.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  displayName?: string;
  withSiteUrl?: boolean;
}

class EntraM365GroupGetCommand extends GraphCommand {
  public get name(): string {
    return commands.M365GROUP_GET;
  }

  public get description(): string {
    return 'Gets information about the specified Microsoft 365 Group or Microsoft Teams team';
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
        withSiteUrl: !!args.options.withSiteUrl
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
        option: '--withSiteUrl'
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
    this.optionSets.push({ options: ['id', 'displayName'] });
  }

  #initTypes(): void {
    this.types.string.push('id', 'displayName');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let group: GroupExtended;

    try {
      if (args.options.id) {
        group = await entraGroup.getGroupById(args.options.id);
      }
      else {
        group = await entraGroup.getGroupByDisplayName(args.options.displayName!);
      }

      const isUnifiedGroup = await entraGroup.isUnifiedGroup(group.id!);

      if (!isUnifiedGroup) {
        throw Error(`Specified group with id '${group.id}' is not a Microsoft 365 group.`);
      }

      const requestExtendedOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/groups/${group.id}?$select=allowExternalSenders,autoSubscribeNewMembers,hideFromAddressLists,hideFromOutlookClients,isSubscribedByMail`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };
      const groupExtended = await request.get<{ allowExternalSenders: boolean, autoSubscribeNewMembers: boolean, hideFromAddressLists: boolean, hideFromOutlookClients: boolean, isSubscribedByMail: boolean }>(requestExtendedOptions);
      group = { ...group, ...groupExtended };

      if (args.options.withSiteUrl) {
        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/groups/${group.id}/drive?$select=webUrl`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        const res = await request.get<{ webUrl: string }>(requestOptions);
        group.siteUrl = res.webUrl ? res.webUrl.substring(0, res.webUrl.lastIndexOf('/')) : '';
      }

      await logger.log(group);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraM365GroupGetCommand();
