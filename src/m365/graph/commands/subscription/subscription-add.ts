import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

const allowedTlsVersions = ['v1_0', 'v1_1', 'v1_2', 'v1_3'] as const;

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  resource: z.string().alias('r'),
  notificationUrl: z.string().alias('u'),
  changeTypes: z.string().alias('c'),
  expirationDateTime: z.string().optional().alias('e'),
  clientState: z.string().optional().alias('s'),
  lifecycleNotificationUrl: z.string().optional(),
  notificationUrlAppId: z.string().optional(),
  latestTLSVersion: z.enum(allowedTlsVersions).optional(),
  withResourceData: z.boolean().optional(),
  encryptionCertificate: z.string().optional(),
  encryptionCertificateId: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
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
  public get name(): string {
    return commands.SUBSCRIPTION_ADD;
  }

  public get description(): string {
    return 'Creates a Microsoft Graph subscription';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => {
        const url = options.notificationUrl.toLowerCase();
        return url.startsWith('https://') || url.startsWith('eventhub:https://') || url.startsWith('eventgrid:?azuresubscriptionid=');
      }, {
        error: e => `The specified notification URL '${(e.input as Options).notificationUrl}' does not start with either 'https://' or 'EventHub:https://' or 'EventGrid:?azuresubscriptionid='`,
        path: ['notificationUrl']
      })
      .refine(options => this.isValidChangeTypes(options.changeTypes), {
        error: `The specified changeTypes are invalid. Valid options are 'created', 'updated' and 'deleted'`,
        path: ['changeTypes']
      })
      .refine(options => !options.expirationDateTime || validation.isValidISODateTime(options.expirationDateTime), {
        error: 'The expirationDateTime is not a valid ISO date string',
        path: ['expirationDateTime']
      })
      .refine(options => !options.clientState || options.clientState.length <= 128, {
        error: 'The clientState value exceeds the maximum length of 128 characters',
        path: ['clientState']
      })
      .refine(options => {
        if (!options.lifecycleNotificationUrl) {
          return true;
        }
        const url = options.lifecycleNotificationUrl.toLowerCase();
        return url.startsWith('https://') || url.startsWith('eventhub:https://') || url.startsWith('eventgrid:?azuresubscriptionid=');
      }, {
        error: e => `The lifecycle notification URL '${(e.input as Options).lifecycleNotificationUrl}' does not start with either 'https://' or 'EventHub:https://' or 'EventGrid:?azuresubscriptionid='`,
        path: ['lifecycleNotificationUrl']
      })
      .refine(options => !options.withResourceData || options.encryptionCertificate, {
        error: `The 'encryptionCertificate' options is required to include the changed resource data`,
        path: ['encryptionCertificate']
      })
      .refine(options => !options.withResourceData || options.encryptionCertificateId, {
        error: `The 'encryptionCertificateId' options is required to include the changed resource data`,
        path: ['encryptionCertificateId']
      })
      .refine(options => !options.notificationUrlAppId || validation.isValidGuid(options.notificationUrlAppId), {
        error: e => `${(e.input as Options).notificationUrlAppId} is not a valid GUID for the 'notificationUrlAppId'`,
        path: ['notificationUrlAppId']
      });
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