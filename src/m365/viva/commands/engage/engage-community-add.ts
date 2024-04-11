import GlobalOptions from '../../../../GlobalOptions.js';
import { Logger } from '../../../../cli/Logger.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  displayName: string;
  description: string;
  privacy: string;
  adminEntraIds?: string;
  adminEntraUserNames?: string;
}

class VivaEngageCommunityAddCommand extends GraphCommand {
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
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        displayName: args.options.displayName !== 'undefined',
        description: typeof args.options.description !== 'undefined',
        privacy: typeof args.options.privacy !== 'undefined',
        adminEntraIds: typeof args.options.adminEntraIds !== 'undefined',
        adminEntraUserNames: typeof args.options.adminEntraUserNames !== 'undefined'
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
        if (args.options.displayName && args.options.displayName.length > 255) {
          return `The maximum amount of characters for 'displayName' is 255.`;
        }

        if (args.options.description && args.options.description.length > 1024) {
          return `The maximum amount of characters for 'description' is 1024.`;
        }

        if (args.options.privacy && this.privacyOptions.indexOf(args.options.privacy) === -1) {
          return `'${args.options.privacy}' is not a valid value for privacy. Allowed values are: ${this.privacyOptions.join(', ')}.`;
        }

        if (args.options.adminEntraIds && args.options.adminEntraUserNames) {
          return `You can only specify either 'adminEntraIds' or 'adminEntraUserNames'`;
        }

        if (args.options.adminEntraIds) {
          const adminEntraIds = args.options.adminEntraIds.split(',');
          for (let i: number = 0; i < adminEntraIds.length; i++) {
            const trimmedId: string = adminEntraIds[i].trim();
            if (!validation.isValidGuid(trimmedId)) {
              return `${trimmedId} is not a valid GUID`;
            }
          }
        }

        if (args.options.adminEntraUserNames) {
          const adminEntraUserNames = args.options.adminEntraUserNames.split(',');
          for (let i: number = 0; i < adminEntraUserNames.length; i++) {
            const trimmedUPN: string = adminEntraUserNames[i].trim();
            if (!validation.isValidUserPrincipalName(trimmedUPN)) {
              return `${trimmedUPN} is not a valid UPN`;
            }
          }
        }

        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('displayName', 'description', 'privacy', 'adminEntraIds', 'adminEntraUserNames');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const { displayName, description, privacy, adminEntraIds, adminEntraUserNames, wait } = args.options;

    if (this.verbose) {
      logger.logToStderr(`Creating a Viva Engage community with display name '${displayName}'...`);
    }

    try {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/beta/employeeExperience/communities`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json;odata=nometadata'
        },
        responseType: 'json',
        fullResponse: true,
        data: {
          displayName: displayName,
          description: description,
          privacy: privacy
        }
      };

      let entraIds: string[] = [];

      if (adminEntraIds) {
        entraIds = adminEntraIds.split(',').map(id => this.getUserIdUrl(id));
      }
      else if (adminEntraUserNames) {
        const userUPNs = adminEntraUserNames.split(',').map(upn => this.getUserId(upn));
        const userIds = await Promise.all(userUPNs);
        entraIds = userIds.map(id => this.getUserIdUrl(id));
      }

      if (entraIds.length > 0) {
        requestOptions.data['owners@odata.bind'] = entraIds;
      }

      const res: any = await request.post(requestOptions);

      const location: string = res.headers.location;

      if (!wait) {
        await logger.log(location);
        return;
      }

      let status: string;

      do {
        if (this.verbose) {
          logger.logToStderr(`Waiting 30 seconds...`);
        }

        await new Promise(resolve => setTimeout(resolve, 30000));

        if (this.debug) {
          logger.logToStderr(`Checking create community operation status...`);
        }

        const operation: any = await request.get({
          url: location,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        });
        status = operation.status;

        if (this.verbose) {
          logger.logToStderr(`Community creation operation status: ${status}`);
        }

        if (status === 'failed') {
          throw `Community creation failed: ${operation.error?.message}`;
        }

        if (status === 'succeeded') {
          await logger.log(operation);
        }
      }
      while (status === 'inprogress');
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getUserId(userUPN: string): Promise<string> {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/users/${userUPN}?$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const res = await request.get<{ id: string }>(requestOptions);
    return res.id;
  }

  private getUserIdUrl(id: string): string {
    return `https://graph.microsoft.com/beta/users/${id}`;
  }
}

export default new VivaEngageCommunityAddCommand();