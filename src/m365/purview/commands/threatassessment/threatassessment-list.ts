import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

const allowedTypes = ['mail', 'file', 'emailFile', 'url'] as const;

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  type: z.enum(allowedTypes).optional().alias('t')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PurviewThreatAssessmentListCommand extends GraphCommand {
  public get name(): string {
    return commands.THREATASSESSMENT_LIST;
  }

  public get description(): string {
    return 'Gets a list of threat assessments';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'type', 'category'];
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