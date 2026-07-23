import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  taskId: z.string().alias('i'),
  url: z.string().alias('u'),
  alias: z.string().optional(),
  type: z.string()
    .refine(val => ['powerpoint', 'word', 'excel', 'other'].includes(val.toLocaleLowerCase()), {
      message: 'The type must be one of the following: PowerPoint, Word, Excel, Other.'
    })
    .optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PlannerTaskReferenceAddCommand extends GraphCommand {
  public get name(): string {
    return commands.TASK_REFERENCE_ADD;
  }

  public get description(): string {
    return 'Adds a new reference to a Planner task';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const etag = await this.getTaskDetailsEtag(args.options.taskId);
      const requestOptionsTaskDetails: CliRequestOptions = {
        url: `${this.resource}/v1.0/planner/tasks/${formatting.encodeQueryParameter(args.options.taskId)}/details`,
        headers: {
          'accept': 'application/json;odata.metadata=none',
          'If-Match': etag,
          'Prefer': 'return=representation'
        },
        responseType: 'json',
        data: {
          references: {
            [formatting.openTypesEncoder(args.options.url)]: {
              '@odata.type': 'microsoft.graph.plannerExternalReference',
              previewPriority: ' !',
              ...(args.options.alias && { alias: args.options.alias }),
              ...(args.options.type && { type: args.options.type })
            }
          }
        }
      };
      const res = await request.patch<any>(requestOptionsTaskDetails);
      await logger.log(res.references);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getTaskDetailsEtag(taskId: string): Promise<string> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/planner/tasks/${formatting.encodeQueryParameter(taskId)}/details`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    const response = await request.get<any>(requestOptions);
    return response['@odata.etag'];
  }
}

export default new PlannerTaskReferenceAddCommand();
