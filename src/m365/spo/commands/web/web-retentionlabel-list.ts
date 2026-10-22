import { z } from 'zod';
import { Logger } from "../../../../cli/Logger.js";
import { globalOptionsZod } from '../../../../Command.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from "../../../../utils/odata.js";
import { validation } from "../../../../utils/validation.js";
import SpoCommand from "../../../base/SpoCommand.js";
import commands from "../../commands.js";

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
    error: e => `${e.input} is not a valid SharePoint Online site URL.`
  }).alias('u')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoWebRetentionLabelListCommand extends SpoCommand {
  public get name(): string {
    return commands.WEB_RETENTIONLABEL_LIST;
  }

  public get description(): string {
    return 'Gets a list of retention labels that are available on a site';
  }

  public defaultProperties(): string[] | undefined {
    return ['TagId', 'TagName'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving all retention labels that are available on ${args.options.webUrl}...`);
    }

    const requestUrl: string = `${args.options.webUrl}/_api/SP.CompliancePolicy.SPPolicyStoreProxy.GetAvailableTagsForSite(siteUrl=@a1)?@a1='${formatting.encodeQueryParameter(args.options.webUrl)}'`;

    try {
      const response = await odata.getAllItems(requestUrl);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoWebRetentionLabelListCommand();