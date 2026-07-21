import { PlannerTaskDetails } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  url: z.string()
    .refine(val => val.indexOf('https://') === 0 || val.indexOf('http://') === 0, {
      message: 'The url option should contain a valid URL. A valid URL starts with http(s)://'
    })
    .optional()
    .alias('u'),
  alias: z.string().optional(),
  taskId: z.string().alias('i'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PlannerTaskReferenceRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.TASK_REFERENCE_REMOVE;
  }

  public get description(): string {
    return 'Removes the reference from the Planner task';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodType | undefined {
    return schema
      .refine(opts => [opts.url, opts.alias].filter(x => x !== undefined).length === 1, {
        message: `Specify exactly one of the following options: 'url' or 'alias'.`,
        params: {
          customCode: 'optionSet',
          options: ['url', 'alias']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removeReference(logger, args);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the reference from the Planner task?` });

      if (result) {
        await this.removeReference(logger, args);
      }
    }
  }

  private async removeReference(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const { etag, url } = await this.getTaskDetailsEtagAndUrl(args.options);
      const requestOptionsTaskDetails: CliRequestOptions = {
        url: `${this.resource}/v1.0/planner/tasks/${args.options.taskId}/details`,
        headers: {
          'accept': 'application/json;odata.metadata=none',
          'If-Match': etag,
          'Prefer': 'return=representation'
        },
        responseType: 'json',
        data: {
          references: {
            [formatting.openTypesEncoder(url)]: null
          }
        }
      };

      await request.patch(requestOptionsTaskDetails);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getTaskDetailsEtagAndUrl(options: Options): Promise<{ etag: string, url: string }> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/planner/tasks/${formatting.encodeQueryParameter(options.taskId)}/details`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    let url: string = options.url!;

    const taskDetails = await request.get<PlannerTaskDetails>(requestOptions);
    if (options.alias) {
      const urls: string[] = [];

      if (taskDetails.references) {
        Object.entries(taskDetails.references!).forEach((ref: any) => {
          if (ref[1].alias?.toLocaleLowerCase() === options.alias!.toLocaleLowerCase()) {
            urls.push(decodeURIComponent(ref[0]));
          }
        });
      }

      if (urls.length === 0) {
        throw `The specified reference with alias ${options.alias} does not exist`;
      }

      if (urls.length > 1) {
        throw `Multiple references with alias ${options.alias} found. Pass one of the following urls within the "--url" option : ${urls}`;
      }

      url = urls[0];
    }

    return { etag: (taskDetails as any)['@odata.etag'], url };
  }
}

export default new PlannerTaskReferenceRemoveCommand();
