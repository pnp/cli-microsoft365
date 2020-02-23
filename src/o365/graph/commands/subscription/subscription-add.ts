import commands from '../../commands';
import request from '../../../../request';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import GraphCommand from '../../../base/GraphCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

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
    return `${commands.SUBSCRIPTION_ADD}`;
  }

  public get description(): string {
    return 'Creates a Microsoft Graph subscription';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.changeType = args.options.changeType;
    telemetryProps.expirationDateTime = typeof args.options.expirationDateTime !== 'undefined';
    telemetryProps.clientState = typeof args.options.clientState !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const body: any = {
      changeType: args.options.changeType,
      resource: args.options.resource,
      notificationUrl: args.options.notificationUrl,
      expirationDateTime: this.getExpirationDateTimeOrDefault(cmd, args)
    };

    if (args.options.clientState) {
      body["clientState"] = args.options.clientState;
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/subscriptions`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json'
      },
      body,
      json: true
    };

    request
      .post(requestOptions)
      .then((res: any): void => {
        cmd.log(res);

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  private getExpirationDateTimeOrDefault(cmd: CommandInstance, args: CommandArgs): string {
    if (args.options.expirationDateTime) {
      if (this.debug) {
        cmd.log(`Expiration date time is specified (${args.options.expirationDateTime}).`);
      }

      return args.options.expirationDateTime;
    }

    if (this.debug) {
      cmd.log(`Expiration date time is not specified. Will try to get appropriate maximum value`);
    }

    const fromNow = (minutes: number) => {
      // convert minutes in milliseconds
      return new Date(Date.now() + (minutes * 60000));
    }

    const expirationDelayPerResource: any = DEFAULT_EXPIRATION_DELAY_IN_MINUTES_PER_RESOURCE_TYPE;

    for (let resource in expirationDelayPerResource) {
      if (args.options.resource.indexOf(resource) < 0) {
        continue;
      }

      const resolvedExpirationDelay = expirationDelayPerResource[resource] as number;

      // Compute the actual expirationDateTime argument from now
      const actualExpiration = fromNow(resolvedExpirationDelay - SAFE_MINUTES_DELTA);
      const actualExpirationIsoString = actualExpiration.toISOString();

      if (this.debug) {
        cmd.log(`Matching resource in default values '${args.options.resource}' => '${resource}'`);
        cmd.log(`Resolved expiration delay: ${resolvedExpirationDelay} (safe delta: ${SAFE_MINUTES_DELTA})`);
        cmd.log(`Actual expiration date time: ${actualExpirationIsoString}`);
      }

      if (this.verbose) {
        cmd.log(`An expiration maximum delay is resolved for the resource '${args.options.resource}' : ${resolvedExpirationDelay} minutes.`);
      }

      return actualExpirationIsoString;
    }

    // If an resource specific expiration has not been found, return a default expiration delay
    if (this.verbose) {
      cmd.log(`An expiration maximum delay couldn't be resolved for the resource '${args.options.resource}'. Will use generic default value: ${DEFAULT_EXPIRATION_DELAY_IN_MINUTES} minutes.`);
    }

    const actualExpiration = fromNow(DEFAULT_EXPIRATION_DELAY_IN_MINUTES - SAFE_MINUTES_DELTA);
    const actualExpirationIsoString = actualExpiration.toISOString();

    if (this.debug) {
      cmd.log(`Actual expiration date time: ${actualExpirationIsoString}`);
    }

    return actualExpirationIsoString;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-r, --resource <resource>',
        description: `The resource that will be monitored for changes`
      },
      {
        option: '-u, --notificationUrl <notificationUrl>',
        description: 'The URL of the endpoint that will receive the notifications. This URL must use the HTTPS protocol'
      },
      {
        option: '-c, --changeType <changeType>',
        description: `The type of change in the subscribed resource that will raise a notification. The supported values are: created, updated, deleted. Multiple values can be combined using a comma-separated list`,
        autocomplete: ['created', 'updated', 'deleted']
      },
      {
        option: '-e, --expirationDateTime [expirationDateTime]',
        description: `The date and time when the webhook subscription expires. The time is in UTC, and can be an amount of time from subscription creation that varies for the resource subscribed to. If not specified, the maximum allowed expiration for the specified resource will be used`
      },
      {
        option: '-s, --clientState [clientState]',
        description: `The value of the clientState property sent by the service in each notification. The maximum length is 128 characters`
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.resource) {
        return 'Required option resource is missing';
      }

      if (!args.options.notificationUrl) {
        return 'Required option notificationUrl is missing';
      }

      if (args.options.notificationUrl.indexOf('https://') !== 0) {
        return `The specified notification URL '${args.options.notificationUrl}' does not start with 'https://'`;
      }

      if (!args.options.changeType) {
        return 'Required option changeType is missing';
      }

      if (!this.isValidChangeTypes(args.options.changeType)) {
        return `The specified changeType is invalid. Valid options are 'created', 'updated' and 'deleted'`;
      }

      if (args.options.expirationDateTime && !Utils.isValidISODateTime(args.options.expirationDateTime)) {
        return 'The expirationDateTime is not a valid ISO date string';
      }

      if (args.options.clientState && args.options.clientState.length > 128) {
        return 'The clientState value exceeds the maximum length of 128 characters';
      }

      return true;
    };
  }


  private isValidChangeTypes(changeTypes: string): boolean {
    const validChangeTypes = ["created", "updated", "deleted"];
    const invalidChangesTypes = changeTypes.split(",").filter(c => validChangeTypes.indexOf(c.trim()) < 0);

    return invalidChangesTypes.length === 0;
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    On personal OneDrive, you can subscribe to the root folder or any subfolder
    in that drive. On OneDrive for Business, you can subscribe to only the root
    folder.

    Notifications are sent for the requested types of changes on the subscribed
    folder, or any file, folder, or other ${chalk.grey(`driveItem`)} instances in its hierarchy.
    You cannot subscribe to ${chalk.grey(`drive`)} or ${chalk.grey(`driveItem`)} instances that are not folders,
    such as individual files.

    In Outlook, delegated permission supports subscribing to items in folders in
    only the signed-in user's mailbox. That means, for example, you cannot use
    the delegated permission Calendars.Read to subscribe to events in another
    user’s mailbox.

    To subscribe to change notifications of Outlook contacts, events,
    or messages in shared or delegated folders:

    - Use the corresponding application permission to subscribe to changes of
      items in a folder or mailbox of any user in the tenant.
    - Do not use the Outlook sharing permissions (Contacts.Read.Shared,
      Calendars.Read.Shared, Mail.Read.Shared, and their read/write
      counterparts), as they do not support subscribing to change notifications
      on items in shared or delegated folders.

  Examples:
  
    Create a subscription
      ${this.name} --resource "me/mailFolders('Inbox')/messages" --changeType "updated" --notificationUrl "https://webhook.azurewebsites.net/api/send/myNotifyClient" --expirationDateTime "2016-11-20T18:23:45.935Z" --clientState "secretClientState"

    Create a subscription on multiple change types
      ${this.name} --resource groups --changeType updated,deleted --notificationUrl "https://webhook.azurewebsites.net/api/send/myNotifyClient" --expirationDateTime "2016-11-20T18:23:45.935Z" --clientState "secretClientState"

    Create a subscription using the maximum allowed expiration for Group resources
      ${this.name} --resource groups --changeType "updated" --notificationUrl "https://webhook.azurewebsites.net/api/send/myNotifyClient"
`);
  }
}

module.exports = new GraphSubscriptionAddCommand();