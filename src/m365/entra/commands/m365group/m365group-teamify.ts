import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  displayName?: string;
  mailNickname?: string;
}

class EntraM365GroupTeamifyCommand extends GraphCommand {
  public get name(): string {
    return commands.M365GROUP_TEAMIFY;
  }

  public get description(): string {
    return 'Creates a new Microsoft Teams team under existing Microsoft 365 group';
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
        id: typeof args.options.id !== 'undefined',
        mailNickname: typeof args.options.mailNickname !== 'undefined',
        displayName: typeof args.options.displayName !== 'undefined'
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
        option: '--mailNickname [mailNickname]'
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
    this.optionSets.push({ options: ['id', 'dispalyName', 'mailNickname'] });
  }

  #initTypes(): void {
    this.types.string.push('id', 'displayName', 'mailNickname');
  }

  private async getGroupId(options: Options): Promise<string> {
    if (options.id) {
      return options.id;
    }

    if (options.displayName) {
      return await entraGroup.getGroupIdByDisplayName(options.displayName);
    }

    return await entraGroup.getGroupIdByMailNickname(options.mailNickname!);
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const groupId = await this.getGroupId(args.options);
      const isUnifiedGroup = await entraGroup.isUnifiedGroup(groupId);

      if (!isUnifiedGroup) {
        throw Error(`Specified group with id '${groupId}' is not a Microsoft 365 group.`);
      }

      const data: any = {
        "memberSettings": {
          "allowCreatePrivateChannels": true,
          "allowCreateUpdateChannels": true
        },
        "messagingSettings": {
          "allowUserEditMessages": true,
          "allowUserDeleteMessages": true
        },
        "funSettings": {
          "allowGiphy": true,
          "giphyContentRating": "strict"
        }
      };


      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/groups/${formatting.encodeQueryParameter(groupId)}/team`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        data: data,
        responseType: 'json'
      };

      await request.put(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraM365GroupTeamifyCommand();