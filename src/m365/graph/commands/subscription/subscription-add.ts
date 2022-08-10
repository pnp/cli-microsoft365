import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  resource: string;
  changeType: string;
  notificationUrl: string;
  expirationDateTime?: string;
  clientState?: string;
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

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        changeType: args.options.changeType,
        expirationDateTime: typeof args.options.expirationDateTime !== 'undefined',
        clientState: typeof args.options.clientState !== 'undefined'
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
        option: '-c, --changeType <changeType>',
        autocomplete: ['created', 'updated', 'deleted']
      },
      {
        option: '-e, --expirationDateTime [expirationDateTime]'
      },
      {
        option: '-s, --clientState [clientState]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.notificationUrl.indexOf('https://') !== 0) {
          return `The specified notification URL '${args.options.notificationUrl}' does not start with 'https://'`;
        }
    
        if (!this.isValidChangeTypes(args.options.changeType)) {
          return `The specified changeType is invalid. Valid options are 'created', 'updated' and 'deleted'`;
        }
    
        if (args.options.expirationDateTime && !validation.isValidISODateTime(args.options.expirationDateTime)) {
          return 'The expirationDateTime is not a valid ISO date string';
        }
    
        if (args.options.clientState && args.options.clientState.length > 128) {
          return 'The clientState value exceeds the maximum length of 128 characters';
        }
    
        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const data: any = {
      changeType: args.options.changeType,
      resource: args.options.resource,
      notificationUrl: args.options.notificationUrl,
      expirationDateTime: this.getExpirationDateTimeOrDefault(logger, args)
    };

    if (args.options.clientState) {
      data["clientState"] = args.options.clientState;
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/subscriptions`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json'
      },
      data,
      responseType: 'json'
    };

    request
      .post(requestOptions)
      .then((res: any): void => {
        logger.log(res);
        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getExpirationDateTimeOrDefault(logger: Logger, args: CommandArgs): string {
    if (args.options.expirationDateTime) {
      if (this.debug) {
        logger.logToStderr(`Expiration date time is specified (${args.options.expirationDateTime}).`);
      }

      return args.options.expirationDateTime;
    }

    if (this.debug) {
      logger.logToStderr(`Expiration date time is not specified. Will try to get appropriate maximum value`);
    }

    const fromNow = (minutes: number) => {
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
        logger.logToStderr(`Matching resource in default values '${args.options.resource}' => '${resource}'`);
        logger.logToStderr(`Resolved expiration delay: ${resolvedExpirationDelay} (safe delta: ${SAFE_MINUTES_DELTA})`);
        logger.logToStderr(`Actual expiration date time: ${actualExpirationIsoString}`);
      }

      if (this.verbose) {
        logger.logToStderr(`An expiration maximum delay is resolved for the resource '${args.options.resource}' : ${resolvedExpirationDelay} minutes.`);
      }

      return actualExpirationIsoString;
    }

    // If an resource specific expiration has not been found, return a default expiration delay
    if (this.verbose) {
      logger.logToStderr(`An expiration maximum delay couldn't be resolved for the resource '${args.options.resource}'. Will use generic default value: ${DEFAULT_EXPIRATION_DELAY_IN_MINUTES} minutes.`);
    }

    const actualExpiration = fromNow(DEFAULT_EXPIRATION_DELAY_IN_MINUTES - SAFE_MINUTES_DELTA);
    const actualExpirationIsoString = actualExpiration.toISOString();

    if (this.debug) {
      logger.logToStderr(`Actual expiration date time: ${actualExpirationIsoString}`);
    }

    return actualExpirationIsoString;
  }

  private isValidChangeTypes(changeTypes: string): boolean {
    const validChangeTypes = ["created", "updated", "deleted"];
    const invalidChangesTypes = changeTypes.split(",").filter(c => validChangeTypes.indexOf(c.trim()) < 0);

    return invalidChangesTypes.length === 0;
  }
}

module.exports = new GraphSubscriptionAddCommand();