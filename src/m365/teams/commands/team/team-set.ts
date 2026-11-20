import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import { zod } from '../../../../utils/zod.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

const VisibilityEnum = {
  Private: 'Private',
  Public: 'Public'
} as const;

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value must be a valid GUID.'
    })
    .optional()
    .alias('i'),
  name: z.string().optional().alias('n'),
  description: z.string().optional(),
  mailNickName: z.string().optional(),
  classification: z.string().optional(),
  visibility: zod.coercedEnum(VisibilityEnum).optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsTeamSetCommand extends GraphCommand {
  public get name(): string {
    return commands.TEAM_SET;
  }

  public get description(): string {
    return 'Updates settings of a Microsoft Teams team';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  private mapRequestBody(options: Options): any {
    const requestBody: any = {};
    if (options.name) {
      requestBody.displayName = options.name;
    }
    if (options.description) {
      requestBody.description = options.description;
    }
    if (options.mailNickName) {
      requestBody.mailNickName = options.mailNickName;
    }
    if (options.classification) {
      requestBody.classification = options.classification;
    }
    if (options.visibility) {
      requestBody.visibility = options.visibility;
    }
    return requestBody;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const data: any = this.mapRequestBody(args.options);

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/groups/${formatting.encodeQueryParameter(args.options.id as string)}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      data: data,
      responseType: 'json'
    };

    try {
      await request.patch(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TeamsTeamSetCommand();