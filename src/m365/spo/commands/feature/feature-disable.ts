import { z } from 'zod';
import { Logger } from "../../../../cli/Logger.js";
import { globalOptionsZod } from "../../../../Command.js";
import request, { CliRequestOptions } from "../../../../request.js";
import { validation } from '../../../../utils/validation.js';
import { zod } from '../../../../utils/zod.js';
import SpoCommand from "../../../base/SpoCommand.js";
import commands from "../../commands.js";

enum FeatureScope {
  Site = 'Site',
  Web = 'Web'
}

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string().alias('u'),
  id: z.string().refine(id => validation.isValidGuid(id), {
    error: e => `${e.input} is not a valid GUID for option id.`
  }).alias('i'),
  scope: zod.coercedEnum(FeatureScope).optional().alias('s'),
  force: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFeatureDisableCommand extends SpoCommand {
  public get name(): string {
    return commands.FEATURE_DISABLE;
  }

  public get description(): string {
    return 'Disables feature for the specified site or web';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const scope: string = args.options.scope ?? 'Web';
    const force: boolean = !!args.options.force;

    if (this.verbose) {
      await logger.logToStderr(`Disabling feature '${args.options.id}' on scope '${scope}' for url '${args.options.webUrl}' (force='${force}')...`);
    }

    const url: string = `${args.options.webUrl}/_api/${scope}/features/remove(featureId=guid'${args.options.id}',force=${force})`;

    const requestOptions: CliRequestOptions = {
      url: url,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoFeatureDisableCommand();