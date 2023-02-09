import { AxiosRequestConfig } from 'axios';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

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
}

class PurviewRetentionLabelAddCommand extends GraphCommand {
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
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        retentionTrigger: typeof args.options.retentionTrigger !== 'undefined',
        defaultRecordBehavior: typeof args.options.defaultRecordBehavior !== 'undefined',
        descriptionForUsers: typeof args.options.descriptionForUsers !== 'undefined',
        descriptionForAdmins: typeof args.options.descriptionForAdmins !== 'undefined',
        labelToBeApplied: typeof args.options.labelToBeApplied !== 'undefined'
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
        autocomplete: ['doNotRetain', 'retain', 'retainAsRecord', 'retainAsRegulatoryRecord']
      },
      {
        option: '--actionAfterRetentionPeriod <actionAfterRetentionPeriod>',
        autocomplete: ['none', 'delete', 'startDispositionReview']
      },
      {
        option: '--retentionDuration <retentionDuration>'
      },
      {
        option: '-t, --retentionTrigger [retentionTrigger]',
        autocomplete: ['dateLabeled', 'dateCreated', 'dateModified', 'dateOfEvent']
      },
      {
        option: '--defaultRecordBehavior [defaultRecordBehavior]',
        autocomplete: ['startLocked', 'startUnlocked']
      },
      {
        option: '--descriptionForUsers [descriptionForUsers]'
      },
      {
        option: '--descriptionForAdmins [descriptionForAdmins]'
      },
      {
        option: '--labelToBeApplied [labelToBeApplied]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (isNaN(args.options.retentionDuration)) {
          return `Specified retentionDuration ${args.options.retentionDuration} is not a number`;
        }

        if (['doNotRetain', 'retain', 'retainAsRecord', 'retainAsRegulatoryRecord'].indexOf(args.options.behaviorDuringRetentionPeriod) === -1) {
          return `${args.options.behaviorDuringRetentionPeriod} is not a valid behavior of a document with the label. Allowed values are doNotRetain|retain|retainAsRecord|retainAsRegulatoryRecord`;
        }

        if (['none', 'delete', 'startDispositionReview'].indexOf(args.options.actionAfterRetentionPeriod) === -1) {
          return `${args.options.actionAfterRetentionPeriod} is not a valid action to take on a document with the label. Allowed values are none|delete|startDispositionReview`;
        }

        if (args.options.retentionTrigger &&
          ['dateLabeled', 'dateCreated', 'dateModified', 'dateOfEvent'].indexOf(args.options.retentionTrigger) === -1) {
          return `${args.options.retentionTrigger} is not a valid action retention duration calculation. Allowed values are dateLabeled|dateCreated|dateModified|dateOfEvent`;
        }

        if (args.options.defaultRecordBehavior &&
          ['startLocked', 'startUnlocked'].indexOf(args.options.defaultRecordBehavior) === -1) {
          return `${args.options.defaultRecordBehavior} is not a valid state of a record label. Allowed values are startLocked|startUnlocked`;
        }

        return true;
      }
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

    const requestOptions: AxiosRequestConfig = {
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
}

module.exports = new PurviewRetentionLabelAddCommand();