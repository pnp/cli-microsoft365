import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils/odata';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  type?: string;
}

class PurviewThreatassesmentCommand extends GraphCommand {
  private static readonly allowedTypes: string[] = ['mail', 'file', 'emailFile', 'url'];

  public get name(): string {
    return commands.THREATASSESMENT_LIST;
  }

  public get description(): string {
    return 'Get a list of threat assessment';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'contentType', 'category'];
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
        option: '-t, --type [type]', autocomplete: PurviewThreatassesmentCommand.allowedTypes
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {

        if (args.options.type && PurviewThreatassesmentCommand.allowedTypes.indexOf(args.options.type) < 0) {
          return `${args.options.type} is not a valid type. Valid types are ${PurviewThreatassesmentCommand.allowedTypes.join(', ')}`;
        }
        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving list of threat assessment`);
    }

    let url = `${this.resource}/v1.0/informationProtection/threatAssessmentRequests`;
    if (args.options.type) {
      url += '';
    }

    try {
      const items = await odata.getAllItems<any>(url, 'minimal');
      if (args.options.type) {
        let type: string;
        switch (args.options.type) {
          case 'mail':
            type = '#microsoft.graph.mailAssessmentRequest';
            break;
          case 'file':
            type = '#microsoft.graph.fileAssessmentRequest';
            break;
          case 'emailFile':
            type = '#microsoft.graph.emailFileAssessmentRequest';
            break;
          case 'url':
            type = '#microsoft.graph.urlAssessmentRequest';
            break;
        }

        const filteredItems = items.filter(item => item['@odata.type'] === type);
        logger.log(filteredItems);
      }
      else {
        logger.log(items);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new PurviewThreatassesmentCommand();