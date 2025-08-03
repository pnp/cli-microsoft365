import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  resource: string;
  changeTypes: string;
  notificationUrl: string;
  expirationDateTime?: string;
  clientState?: string;
  lifecycleNotificationUrl?: string;
  notificationUrlAppId?: string;
  latestTLSVersion?: string;
  withResourceData?: boolean;
  encryptionCertificate?: string;
  encryptionCertificateId?: string;
}

const DEFAULT_EXPIRATION_DELAY_IN_MINUTES_PER_RESOURCE_TYPE = {
  // User, group, other directory resources	4230 minutes (under 3 days)
  "users": 4230,
  "groups": 4230,
  // Mail	4230 minutes (under 3 days)
  "/messages": 4230,
  // Calendar	4230 minutes (under 3 days)
  "/events": 4230,
  // Contacts	4230 minutes (under 3 days)
  "/contacts": 4230,
  // Group conversations	4230 minutes (under 3 days)
  "/conversations": 4230,
  // Drive root items	4230 minutes (under 3 days)
  "/drive/root": 4230,
  // Security alerts	43200 minutes (under 30 days)
  "security/alerts": 43200
};
const DEFAULT_EXPIRATION_DELAY_IN_MINUTES = 4230;
const SAFE_MINUTES_DELTA = 1;

class GraphSubscriptionAddCommand extends GraphCommand {
  private allowedTlsVersions: string[] = ['v1_0', 'v1_1', 'v1_2', 'v1_3'];
  public get name(): string {
    return commands.SUBSCRIPTION_ADD;
  }

  public get description(): string {
    return 'Creates a Microsoft Graph subscription';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        changeTypes: args.options.changeTypes,
        expirationDateTime: typeof args.options.expirationDateTime !== 'undefined',
        clientState: typeof args.options.clientState !== 'undefined',
        lifecycleNotificationUrl: typeof args.options.lifecycleNotificationUrl !== 'undefined',
        notificationUrlAppId: typeof args.options.notificationUrlAppId !== 'undefined',
        latestTLSVersion: typeof args.options.latestTLSVersion !== 'undefined',
        withResourceData: !!args.options.withResourceData,
        encryptionCertificate: typeof args.options.encryptionCertificate !== 'undefined',
        encryptionCertificateId: typeof args.options.encryptionCertificateId !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-r, --resource <resource>'
      },
      {
        option: '-u, --notificationUrl <notificationUrl>'
      },
      {
        option: '-c, --changeTypes <changeTypes>',
        autocomplete: ['created', 'updated', 'deleted']
      },
      {
        option: '-e, --expirationDateTime [expirationDateTime]'
      },
      {
        option: '-s, --clientState [clientState]'
      },
      {
        option: '--lifecycleNotificationUrl [lifecycleNotificationUrl]'
      },
      {
        option: '--notificationUrlAppId [notificationUrlAppId]'
      },
      {
        option: '--latestTLSVersion [latestTLSVersion]',
        autocomplete: this.allowedTlsVersions
      },
      {
        option: '--withResourceData [withResourceData]'
      },
      {
        option: '--encryptionCertificate [encryptionCertificate]'
      },
      {
        option: '--encryptionCertificateId [encryptionCertificateId]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!args.options.notificationUrl.toLowerCase().startsWith('https://')
          && !args.options.notificationUrl.toLowerCase().startsWith('eventhub:https://')
          && !args.options.notificationUrl.toLowerCase().startsWith('eventgrid:?azuresubscriptionid=')) {
          return `The specified notification URL '${args.options.notificationUrl}' does not start with either 'https://' or 'EventHub:https://' or 'EventGrid:?azuresubscriptionid='`;
        }

        if (!this.isValidChangeTypes(args.options.changeTypes)) {
          return `The specified changeTypes are invalid. Valid options are 'created', 'updated' and 'deleted'`;
        }

        if (args.options.expirationDateTime && !validation.isValidISODateTime(args.options.expirationDateTime)) {
          return 'The expirationDateTime is not a valid ISO date string';
        }

        if (args.options.clientState && args.options.clientState.length > 128) {
          return 'The clientState value exceeds the maximum length of 128 characters';
        }

        if (args.options.lifecycleNotificationUrl && !args.options.lifecycleNotificationUrl.toLowerCase().startsWith('https://')
          && !args.options.lifecycleNotificationUrl.toLowerCase().startsWith('eventhub:https://')
          && !args.options.lifecycleNotificationUrl.toLowerCase().startsWith('eventgrid:?azuresubscriptionid=')) {
          return `The lifecycle notification URL '${args.options.lifecycleNotificationUrl}' does not start with either 'https://' or 'EventHub:https://' or 'EventGrid:?azuresubscriptionid='`;
        }

        if (args.options.latestTLSVersion && this.allowedTlsVersions.map(x => x.toLowerCase()).indexOf(args.options.latestTLSVersion.toLowerCase()) === -1) {
          return `${args.options.latestTLSVersion} is not a valid TLS version. Allowed values are ${this.allowedTlsVersions.join(', ')}`;
        }

        if (args.options.withResourceData && !args.options.encryptionCertificate) {
          return `The 'encryptionCertificate' options is required to include the changed resource data`;
        }

        if (args.options.withResourceData && !args.options.encryptionCertificateId) {
          return `The 'encryptionCertificateId' options is required to include the changed resource data`;
        }

        if (args.options.notificationUrlAppId && !validation.isValidGuid(args.options.notificationUrlAppId)) {
          return `${args.options.notificationUrlAppId} is not a valid GUID for the 'notificationUrlAppId'`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const data: any = {
      changeType: args.options.changeTypes,
      resource: args.options.resource,
      notificationUrl: args.options.notificationUrl,
      expirationDateTime: await this.getExpirationDateTimeOrDefault(logger, args),
      clientState: args.options.clientState,
      includeResourceData: args.options.withResourceData,
      encryptionCertificate: args.options.encryptionCertificate,
      encryptionCertificateId: args.options.encryptionCertificateId,
      lifecycleNotificationUrl: args.options.lifecycleNotificationUrl,
      notificationUrlAppId: args.options.notificationUrlAppId,
      latestSupportedTlsVersion: args.options.latestTLSVersion
    };

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/subscriptions`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json'
      },
      data,
      responseType: 'json'
    };

    try {
      const res = await request.post(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getExpirationDateTimeOrDefault(logger: Logger, args: CommandArgs): Promise<string> {
    if (args.options.expirationDateTime) {
      if (this.debug) {
        await logger.logToStderr(`Expiration date time is specified (${args.options.expirationDateTime}).`);
      }

      return args.options.expirationDateTime;
    }

    if (this.debug) {
      await logger.logToStderr(`Expiration date time is not specified. Will try to get appropriate maximum value`);
    }

    const fromNow = (minutes: number): Date => {
      // convert minutes in milliseconds
      return new Date(Date.now() + (minutes * 60000));
    };

    const expirationDelayPerResource: any = DEFAULT_EXPIRATION_DELAY_IN_MINUTES_PER_RESOURCE_TYPE;

    for (const resource in expirationDelayPerResource) {
      if (args.options.resource.indexOf(resource) < 0) {
        continue;
      }

      const resolvedExpirationDelay = expirationDelayPerResource[resource] as number;

      // Compute the actual expirationDateTime argument from now
      const actualExpiration = fromNow(resolvedExpirationDelay - SAFE_MINUTES_DELTA);
      const actualExpirationIsoString = actualExpiration.toISOString();

      if (this.debug) {
        await logger.logToStderr(`Matching resource in default values '${args.options.resource}' => '${resource}'`);
        await logger.logToStderr(`Resolved expiration delay: ${resolvedExpirationDelay} (safe delta: ${SAFE_MINUTES_DELTA})`);
        await logger.logToStderr(`Actual expiration date time: ${actualExpirationIsoString}`);
      }

      if (this.verbose) {
        await logger.logToStderr(`An expiration maximum delay is resolved for the resource '${args.options.resource}' : ${resolvedExpirationDelay} minutes.`);
      }

      return actualExpirationIsoString;
    }

    // If an resource specific expiration has not been found, return a default expiration delay
    if (this.verbose) {
      await logger.logToStderr(`An expiration maximum delay couldn't be resolved for the resource '${args.options.resource}'. Will use generic default value: ${DEFAULT_EXPIRATION_DELAY_IN_MINUTES} minutes.`);
    }

    const actualExpiration = fromNow(DEFAULT_EXPIRATION_DELAY_IN_MINUTES - SAFE_MINUTES_DELTA);
    const actualExpirationIsoString = actualExpiration.toISOString();

    if (this.debug) {
      await logger.logToStderr(`Actual expiration date time: ${actualExpirationIsoString}`);
    }

    return actualExpirationIsoString;
  }

  private isValidChangeTypes(changeTypes: string): boolean {
    const validChangeTypes = ["created", "updated", "deleted"];
    const invalidChangesTypes = changeTypes.split(",").filter(c => validChangeTypes.indexOf(c.trim()) < 0);

    return invalidChangesTypes.length === 0;
  }
}

export default new GraphSubscriptionAddCommand();