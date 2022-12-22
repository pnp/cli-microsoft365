import { AxiosRequestConfig } from 'axios';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { validation } from '../../../../utils/validation';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  behaviorDuringRetentionPeriod?: string;
  actionAfterRetentionPeriod?: string;
  retentionDuration?: number;
  retentionTrigger?: string;
  defaultRecordBehavior?: string;
  descriptionForUsers?: string;
  descriptionForAdmins?: string;
  labelToBeApplied?: string;
}

class PurviewRetentionLabelSetCommand extends GraphCommand {
  public allowedBehaviorDuringRetentionPeriodValues = ['doNotRetain', 'retain', 'retainAsRecord', 'retainAsRegulatoryRecord'];
  public allowedActionAfterRetentionPeriodValues = ['none', 'delete', 'startDispositionReview'];
  public allowedRetentionTriggerValues = ['dateLabeled', 'dateCreated', 'dateModified', 'dateOfEvent'];
  public allowedDefaultRecordBehaviorValues = ['startLocked', 'startUnlocked'];

  public get name(): string {
    return commands.RETENTIONLABEL_SET;
  }

  public get description(): string {
    return 'Update a retention label';
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
        behaviorDuringRetentionPeriod: typeof args.options.behaviorDuringRetentionPeriod !== 'undefined',
        actionAfterRetentionPeriod: typeof args.options.actionAfterRetentionPeriod !== 'undefined',
        retentionDuration: typeof args.options.retentionDuration !== 'undefined',
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
        option: '-i, --id <id>'
      },
      {
        option: '--behaviorDuringRetentionPeriod [behaviorDuringRetentionPeriod]',
        autocomplete: this.allowedBehaviorDuringRetentionPeriodValues
      },
      {
        option: '--actionAfterRetentionPeriod [actionAfterRetentionPeriod]',
        autocomplete: this.allowedActionAfterRetentionPeriodValues
      },
      {
        option: '--retentionDuration [retentionDuration]'
      },
      {
        option: '-t, --retentionTrigger [retentionTrigger]',
        autocomplete: this.allowedRetentionTriggerValues
      },
      {
        option: '--defaultRecordBehavior [defaultRecordBehavior]',
        autocomplete: this.allowedDefaultRecordBehaviorValues
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
        if (!validation.isValidGuid(args.options.id)) {
          return `'${args.options.id}' is not a valid GUID.`;
        }

        const { actionAfterRetentionPeriod, behaviorDuringRetentionPeriod, defaultRecordBehavior, descriptionForAdmins, descriptionForUsers, labelToBeApplied, retentionDuration, retentionTrigger } = args.options;
        if ([actionAfterRetentionPeriod, behaviorDuringRetentionPeriod, defaultRecordBehavior, descriptionForAdmins, descriptionForUsers, labelToBeApplied, retentionDuration, retentionTrigger].every(i => typeof i === 'undefined')) {
          return `Specify at least one property to update.`;
        }

        if (behaviorDuringRetentionPeriod && this.allowedBehaviorDuringRetentionPeriodValues.indexOf(behaviorDuringRetentionPeriod) === -1) {
          return `'${behaviorDuringRetentionPeriod}' is not a valid value for the behaviorDuringRetentionPeriod option. Allowed values are ${this.allowedBehaviorDuringRetentionPeriodValues.join('|')}`;
        }

        if (actionAfterRetentionPeriod && this.allowedActionAfterRetentionPeriodValues.indexOf(actionAfterRetentionPeriod) === -1) {
          return `'${actionAfterRetentionPeriod}' is not a valid value for the actionAfterRetentionPeriod option. Allowed values are ${this.allowedActionAfterRetentionPeriodValues.join('|')}`;
        }

        if (retentionTrigger && this.allowedRetentionTriggerValues.indexOf(retentionTrigger) === -1) {
          return `'${retentionTrigger}' is not a valid value for the retentionTrigger option. Allowed values are ${this.allowedRetentionTriggerValues.join('|')}`;
        }

        if (defaultRecordBehavior && this.allowedDefaultRecordBehaviorValues.indexOf(defaultRecordBehavior) === -1) {
          return `'${defaultRecordBehavior}' is not a valid value for the defaultRecordBehavior option. Allowed values are ${this.allowedDefaultRecordBehaviorValues.join('|')}`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.log(`Starting to update retention label with id ${args.options.id}`);
    }
    const requestBody = this.mapRequestBody(args.options);
    const requestOptions: AxiosRequestConfig = {
      url: `${this.resource}/beta/security/labels/retentionLabels/${args.options.id}`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json',
      data: requestBody
    };

    await request.patch(requestOptions);
  }

  private mapRequestBody(options: Options): any {
    const requestBody: any = {};
    const excludeOptions: string[] = [
      'debug',
      'verbose',
      'output',
      'id',
      'retentionDuration'
    ];
    Object.keys(options).forEach(key => {
      if (excludeOptions.indexOf(key) === -1) {
        requestBody[key] = `${(<any>options)[key]}`;
      }
    });

    if (options.retentionDuration) {
      requestBody['retentionDuration'] = {
        '@odata.type': 'microsoft.graph.security.retentionDurationInDays',
        'days': options.retentionDuration
      };
    }
    return requestBody;
  }
}

module.exports = new PurviewRetentionLabelSetCommand();