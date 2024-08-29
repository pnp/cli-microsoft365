import GlobalOptions from '../../../../GlobalOptions.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { vivaEngage } from '../../../../utils/vivaEngage.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  displayName?: string;
  newDisplayName?: string;
  description?: string;
  privacy?: string;
}

class VivaEngageCommunitySetCommand extends GraphCommand {
  private privacyOptions: string[] = ['public', 'private'];

  public get name(): string {
    return commands.ENGAGE_COMMUNITY_SET;
  }

  public get description(): string {
    return 'Updates an existing Viva Engage community';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: typeof args.options.id !== 'undefined',
        displayName: typeof args.options.displayName !== 'undefined',
        newDisplayName: typeof args.options.newDisplayName !== 'undefined',
        description: typeof args.options.description !== 'undefined',
        privacy: typeof args.options.privacy !== 'undefined'
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
        option: '--newDisplayName [newDisplayName]'
      },
      {
        option: '--description [description]'
      },
      {
        option: '--privacy [privacy]',
        autocomplete: this.privacyOptions
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.newDisplayName && args.options.newDisplayName.length > 255) {
          return `The maximum amount of characters for 'newDisplayName' is 255.`;
        }

        if (args.options.description && args.options.description.length > 1024) {
          return `The maximum amount of characters for 'description' is 1024.`;
        }

        if (args.options.privacy && this.privacyOptions.map(x => x.toLowerCase()).indexOf(args.options.privacy.toLowerCase()) === -1) {
          return `${args.options.privacy} is not a valid privacy. Allowed values are ${this.privacyOptions.join(', ')}`;
        }

        if (!args.options.newDisplayName && !args.options.description && !args.options.privacy) {
          return 'Specify at least newDisplayName, description, or privacy.';
        }

        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('id', 'displayName', 'newDisplayName', 'description', 'privacy');
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'displayName'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {

    let communityId = args.options.id;

    if (args.options.displayName) {
      communityId = await vivaEngage.getCommunityIdByDisplayName(args.options.displayName);
    }

    if (this.verbose) {
      await logger.logToStderr(`Updating Viva Engage community with ID ${communityId}...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/employeeExperience/communities/${communityId}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: {
        description: args.options.description,
        displayName: args.options.newDisplayName,
        privacy: args.options.privacy
      }
    };

    try {
      await request.patch(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new VivaEngageCommunitySetCommand();