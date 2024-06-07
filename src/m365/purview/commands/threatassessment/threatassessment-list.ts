import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  type?: string;
}

class PurviewThreatAssessmentListCommand extends GraphCommand {
  private static readonly allowedTypes: string[] = ['mail', 'file', 'emailFile', 'url'];

  public get name(): string {
    return commands.THREATASSESSMENT_LIST;
  }

  public get description(): string {
    return 'Get a list of threat assessments';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'type', 'category'];
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
        type: typeof args.options.type !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-t, --type [type]',
        autocomplete: PurviewThreatAssessmentListCommand.allowedTypes
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.type && PurviewThreatAssessmentListCommand.allowedTypes.indexOf(args.options.type) < 0) {
          return `${args.options.type} is not a valid type. Allowed values are ${PurviewThreatAssessmentListCommand.allowedTypes.join(', ')}`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Retrieving a list of threat assessments');
    }

    try {
      const filter = this.getFilterQuery(args.options);
      const items = await odata.getAllItems<any>(`${this.resource}/v1.0/informationProtection/threatAssessmentRequests${filter}`, 'minimal');

      let itemsToReturn = [];

      switch (args.options.type) {
        case 'mail':
          itemsToReturn = items.filter(item => item['@odata.type'] === '#microsoft.graph.mailAssessmentRequest');
          break;
        case 'emailFile':
          itemsToReturn = items.filter(item => item['@odata.type'] === '#microsoft.graph.emailFileAssessmentRequest');
          break;
        default:
          itemsToReturn = items;
          break;
      }

      for (const item of itemsToReturn) {
        item['type'] = this.getConvertedType(item['@odata.type']);
        delete item['@odata.type'];
      }

      await logger.log(itemsToReturn);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  // Content type is not equal to type. 
  // Threat assessments of type emailFile have contentType mail as well.
  // This function gets the correct filter URL to be able to query the least amount of data
  private getFilterQuery(options: Options): string {
    if (options.type === undefined) {
      return '';
    }

    if (options.type === 'emailFile') {
      return `?$filter=contentType eq 'mail'`;
    }

    return `?$filter=contentType eq '${options.type}'`;
  }

  private getConvertedType(type: string): string {
    switch (type) {
      case '#microsoft.graph.mailAssessmentRequest':
        return 'mail';
      case '#microsoft.graph.fileAssessmentRequest':
        return 'file';
      case '#microsoft.graph.emailFileAssessmentRequest':
        return 'emailFile';
      case '#microsoft.graph.urlAssessmentRequest':
        return 'url';
      default:
        return 'Unknown';
    }
  }
}

export default new PurviewThreatAssessmentListCommand();