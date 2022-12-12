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
        behaviorDuringRetentionPeriod: !!args.options.behaviorDuringRetentionPeriod,
        actionAfterRetentionPeriod: !!args.options.actionAfterRetentionPeriod,
        retentionDuration: !!args.options.retentionDuration,
        retentionTrigger: !!args.options.retentionTrigger,
        defaultRecordBehavior: !!args.options.defaultRecordBehavior,
        descriptionForUsers: !!args.options.descriptionForUsers,
        descriptionForAdmins: !!args.options.descriptionForAdmins,
        labelToBeApplied: !!args.options.labelToBeApplied
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
        autocomplete: ['doNotRetain', 'retain', 'retainAsRecord', 'retainAsRegulatoryRecord']
      },
      {
        option: '--actionAfterRetentionPeriod [actionAfterRetentionPeriod]',
        autocomplete: ['none', 'delete', 'startDispositionReview']
      },
      {
        option: '--retentionDuration [retentionDuration]'
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
        if (!validation.isValidGuid(args.options.id)) {
          return `'${args.options.id}' is not a valid GUID.`;
        }

        const { actionAfterRetionPeriod, behaviorDuringRetentionPeriod, defaultRecordBehavior, descriptionForAdmins, descriptionForUsers, labelToBeApplied, retentionDuration, retentionTrigger } = args.options;
        if ([actionAfterRetionPeriod, behaviorDuringRetentionPeriod, defaultRecordBehavior, descriptionForAdmins, descriptionForUsers, labelToBeApplied, retentionDuration, retentionTrigger].every(i => typeof i === 'undefined')) {
          return `Specify atleast one property to update.`;
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