import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { odata } from '../../../../utils/odata';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  displayName: string;
  behaviorDuringRetentionPeriod: string;
  actionAfterRetentionPeriod: string;
  retentionDuration: number;
  retentionTrigger?: string;
  defaultRecordBehavior?: string;
  descriptionForUsers?: string;
  descriptionForAdmins?: string;
  labelToBeApplied?: string;
  eventTypeId?: string;
  eventTypeName?: string;
}

class PurviewRetentionLabelAddCommand extends GraphCommand {
  private static readonly behaviorDuringRetentionPeriods = ['doNotRetain', 'retain', 'retainAsRecord', 'retainAsRegulatoryRecord'];
  private static readonly actionAfterRetentionPeriods = ['none', 'delete', 'startDispositionReview'];
  private static readonly retentionTriggers = ['dateLabeled', 'dateCreated', 'dateModified', 'dateOfEvent'];
  private static readonly defaultRecordBehavior = ['startLocked', 'startUnlocked'];

  public get name(): string {
    return commands.RETENTIONLABEL_ADD;
  }

  public get description(): string {
    return 'Create a retention label';
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
        retentionTrigger: typeof args.options.retentionTrigger !== 'undefined',
        defaultRecordBehavior: typeof args.options.defaultRecordBehavior !== 'undefined',
        descriptionForUsers: typeof args.options.descriptionForUsers !== 'undefined',
        descriptionForAdmins: typeof args.options.descriptionForAdmins !== 'undefined',
        labelToBeApplied: typeof args.options.labelToBeApplied !== 'undefined',
        eventTypeId: typeof args.options.eventTypeId !== 'undefined',
        eventTypeName: typeof args.options.eventTypeName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --displayName <displayName>'
      },
      {
        option: '--behaviorDuringRetentionPeriod <behaviorDuringRetentionPeriod>',
        autocomplete: PurviewRetentionLabelAddCommand.behaviorDuringRetentionPeriods
      },
      {
        option: '--actionAfterRetentionPeriod <actionAfterRetentionPeriod>',
        autocomplete: PurviewRetentionLabelAddCommand.actionAfterRetentionPeriods
      },
      {
        option: '--retentionDuration <retentionDuration>'
      },
      {
        option: '-t, --retentionTrigger [retentionTrigger]',
        autocomplete: PurviewRetentionLabelAddCommand.retentionTriggers
      },
      {
        option: '--defaultRecordBehavior [defaultRecordBehavior]',
        autocomplete: PurviewRetentionLabelAddCommand.defaultRecordBehavior
      },
      {
        option: '--descriptionForUsers [descriptionForUsers]'
      },
      {
        option: '--descriptionForAdmins [descriptionForAdmins]'
      },
      {
        option: '--labelToBeApplied [labelToBeApplied]'
      },
      {
        option: '--eventTypeId [eventTypeId]'
      },
      {
        option: '--eventTypeName [eventTypeName]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (isNaN(args.options.retentionDuration)) {
          return `Specified retentionDuration ${args.options.retentionDuration} is not a number`;
        }

        if (PurviewRetentionLabelAddCommand.behaviorDuringRetentionPeriods.indexOf(args.options.behaviorDuringRetentionPeriod) === -1) {
          return `${args.options.behaviorDuringRetentionPeriod} is not a valid behavior of a document with the label. Allowed values are ${PurviewRetentionLabelAddCommand.behaviorDuringRetentionPeriods.join(', ')}`;
        }

        if (PurviewRetentionLabelAddCommand.actionAfterRetentionPeriods.indexOf(args.options.actionAfterRetentionPeriod) === -1) {
          return `${args.options.actionAfterRetentionPeriod} is not a valid action to take on a document with the label. Allowed values are ${PurviewRetentionLabelAddCommand.actionAfterRetentionPeriods.join(', ')}`;
        }

        if (args.options.retentionTrigger &&
          PurviewRetentionLabelAddCommand.retentionTriggers.indexOf(args.options.retentionTrigger) === -1) {
          return `${args.options.retentionTrigger} is not a valid action retention duration calculation. Allowed values are ${PurviewRetentionLabelAddCommand.retentionTriggers.join(', ')}`;
        }

        if (args.options.defaultRecordBehavior &&
          PurviewRetentionLabelAddCommand.defaultRecordBehavior.indexOf(args.options.defaultRecordBehavior) === -1) {
          return `${args.options.defaultRecordBehavior} is not a valid state of a record label. Allowed values are ${PurviewRetentionLabelAddCommand.defaultRecordBehavior.join(', ')}`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['eventTypeId', 'eventTypeName'], runsWhen(args) { return args.options.retentionTrigger === 'dateOfEvent'; } }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const retentionTrigger: string = args.options.retentionTrigger ? args.options.retentionTrigger : 'dateLabeled';
    const defaultRecordBehavior: string = args.options.defaultRecordBehavior ? args.options.defaultRecordBehavior : 'startLocked';

    const requestBody: any = {
      displayName: args.options.displayName,
      behaviorDuringRetentionPeriod: args.options.behaviorDuringRetentionPeriod,
      actionAfterRetentionPeriod: args.options.actionAfterRetentionPeriod,
      retentionTrigger: retentionTrigger,
      retentionDuration: {
        '@odata.type': '#microsoft.graph.security.retentionDurationInDays',
        days: args.options.retentionDuration
      },
      defaultRecordBehavior: defaultRecordBehavior
    };

    if (args.options.retentionTrigger === 'dateOfEvent') {
      const eventTypeId = await this.getEventTypeId(args, logger);
      requestBody['retentionEventType@odata.bind'] = `https://graph.microsoft.com/beta/security/triggerTypes/retentionEventTypes/${eventTypeId}`;
    }

    if (args.options.descriptionForAdmins) {
      if (this.verbose) {
        logger.logToStderr(`Using '${args.options.descriptionForAdmins}' as descriptionForAdmins`);
      }

      requestBody.descriptionForAdmins = args.options.descriptionForAdmins;
    }

    if (args.options.descriptionForUsers) {
      if (this.verbose) {
        logger.logToStderr(`Using '${args.options.descriptionForUsers}' as descriptionForUsers`);
      }

      requestBody.descriptionForUsers = args.options.descriptionForUsers;
    }

    if (args.options.labelToBeApplied) {
      if (this.verbose) {
        logger.logToStderr(`Using '${args.options.labelToBeApplied}' as labelToBeApplied...`);
      }

      requestBody.labelToBeApplied = args.options.labelToBeApplied;
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/beta/security/labels/retentionLabels`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      data: requestBody,
      responseType: 'json'
    };

    try {
      const response = await request.post(requestOptions);
      logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataPromise(err);
    }
  }

  private async getEventTypeId(args: CommandArgs, logger: Logger): Promise<string> {
    if (args.options.eventTypeId) {
      return args.options.eventTypeId;
    }

    if (this.verbose) {
      logger.logToStderr(`Retrieving the event type id for event type ${args.options.eventTypeName}`);
    }

    const eventTypes = await odata.getAllItems(`${this.resource}/beta/security/triggerTypes/retentionEventTypes`);
    const filteredEventTypes: any = eventTypes.filter((eventType: any) => eventType.displayName === args.options.eventTypeName);

    if (filteredEventTypes.length === 0) {
      throw `The specified retention event type '${args.options.eventTypeName}' does not exist.`;
    }

    return filteredEventTypes[0].id;
  }
}

module.exports = new PurviewRetentionLabelAddCommand();