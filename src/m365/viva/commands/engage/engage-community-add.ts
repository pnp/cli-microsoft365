import GlobalOptions from '../../../../GlobalOptions.js';
import { Logger } from '../../../../cli/Logger.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import { accessToken } from '../../../../utils/accessToken.js';
import auth from '../../../../Auth.js';
import { formatting } from '../../../../utils/formatting.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { setTimeout } from 'timers/promises';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  displayName: string;
  description: string;
  privacy: string;
  adminEntraIds?: string;
  adminEntraUserNames?: string;
  wait?: boolean;
}

class VivaEngageCommunityAddCommand extends GraphCommand {
  private pollingInterval: number = 5000;
  private readonly privacyOptions = ['public', 'private'];

  public get name(): string {
    return commands.ENGAGE_COMMUNITY_ADD;
  }

  public get description(): string {
    return 'Creates a new community in Viva Engage';
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
        adminEntraIds: typeof args.options.adminEntraIds !== 'undefined',
        adminEntraUserNames: typeof args.options.adminEntraUserNames !== 'undefined',
        wait: !!args.options.wait
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '--displayName <displayName>' },
      { option: '--description <description>' },
      {
        option: '--privacy <privacy>',
        autocomplete: this.privacyOptions
      },
      { option: '--adminEntraIds [adminEntraIds]' },
      { option: '--adminEntraUserNames [adminEntraUserNames]' },
      { option: '--wait' }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.displayName.length > 255) {
          return `The maximum amount of characters for 'displayName' is 255.`;
        }

        if (args.options.description.length > 1024) {
          return `The maximum amount of characters for 'description' is 1024.`;
        }

        if (this.privacyOptions.indexOf(args.options.privacy) === -1) {
          return `'${args.options.privacy}' is not a valid value for privacy. Allowed values are: ${this.privacyOptions.join(', ')}.`;
        }

        if (args.options.adminEntraIds) {
          const isValidGUIDArrayResult = validation.isValidGuidArray(args.options.adminEntraIds);
          if (isValidGUIDArrayResult !== true) {
            return `The following GUIDs are invalid for the option 'adminEntraIds': ${isValidGUIDArrayResult}.`;
          }
          if (formatting.splitAndTrim(args.options.adminEntraIds).length > 20) {
            return `Maximum of 20 admins allowed. Please reduce the number of users and try again.`;
          }
        }

        if (args.options.adminEntraUserNames) {
          const isValidUPNArrayResult = validation.isValidUserPrincipalNameArray(args.options.adminEntraUserNames);
          if (isValidUPNArrayResult !== true) {
            return `The following user principal names are invalid for the option 'adminEntraUserNames': ${isValidUPNArrayResult}.`;
          }
          if (formatting.splitAndTrim(args.options.adminEntraUserNames).length > 20) {
            return `Maximum of 20 admins allowed. Please reduce the number of users and try again.`;
          }
        }

        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('displayName', 'description', 'privacy', 'adminEntraIds', 'adminEntraUserNames');
    this.types.boolean.push('wait');
  }

  #initOptionSets(): void {
    this.optionSets.push(
      {
        options: ['adminEntraIds', 'adminEntraUserNames'],
        runsWhen: (args) => args.options.adminEntraIds || args.options.adminEntraUserNames
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const { displayName, description, privacy, adminEntraIds, adminEntraUserNames, wait } = args.options;

    const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[Object.keys(auth.connection.accessTokens)[0]].accessToken);
    if (isAppOnlyAccessToken && !adminEntraIds && !adminEntraUserNames) {
      this.handleError(`Specify at least one admin using either adminEntraIds or adminEntraUserNames options when using application permissions.`);
    }

    if (this.verbose) {
      await logger.logToStderr(`Creating a Viva Engage community with display name '${displayName}'...`);
    }

    try {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/beta/employeeExperience/communities`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json'
        },
        responseType: 'json',
        fullResponse: true,
        data: {
          displayName: displayName,
          description: description,
          privacy: privacy
        }
      };

      const entraIds = await this.getGraphUserUrls(args.options);
      if (entraIds.length > 0) {
        requestOptions.data['owners@odata.bind'] = entraIds;
      }

      const res = await request.post<{ headers: { location: string } }>(requestOptions);

      const location = res.headers.location;

      if (!wait) {
        await logger.log(location);
        return;
      }

      let status: string;
      do {
        if (this.verbose) {
          await logger.logToStderr(`Community still provisioning. Retrying in ${this.pollingInterval / 1000} seconds...`);
        }

        await setTimeout(this.pollingInterval);

        if (this.verbose) {
          await logger.logToStderr(`Checking create community operation status...`);
        }

        const operation = await request.get<any>({
          url: location,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        });
        status = operation.status;

        if (this.verbose) {
          await logger.logToStderr(`Community creation operation status: ${status}`);
        }

        if (status === 'failed') {
          throw `Community creation failed: ${operation.statusDetail}`;
        }

        if (status === 'succeeded') {
          await logger.log(operation);
        }
      }
      while (status === 'notStarted' || status === 'running');
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getGraphUserUrls(options: Options): Promise<string[]> {
    let entraIds: string[] = [];

    if (options.adminEntraIds) {
      entraIds = formatting.splitAndTrim(options.adminEntraIds);
    }
    else if (options.adminEntraUserNames) {
      entraIds = await entraUser.getUserIdsByUpns(formatting.splitAndTrim(options.adminEntraUserNames));
    }

    const graphUserUrls = entraIds.map(id => `${this.resource}/beta/users/${id}`);
    return graphUserUrls;
  }
}

export default new VivaEngageCommunityAddCommand();