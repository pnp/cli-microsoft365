import { Logger } from "../../../../cli/Logger.js";
import GlobalOptions from "../../../../GlobalOptions.js";
import { formatting } from '../../../../utils/formatting.js';
import { odata } from "../../../../utils/odata.js";
import { validation } from "../../../../utils/validation.js";
import SpoCommand from "../../../base/SpoCommand.js";
import commands from "../../commands.js";

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  webUrl: string;
}

class SpoWebRetentionLabelListCommand extends SpoCommand {
  public get name(): string {
    return commands.WEB_RETENTIONLABEL_LIST;
  }

  public get description(): string {
    return `Get a list of retention labels that are available on a site.`;
  }

  public defaultProperties(): string[] | undefined {
    return ['TagId', 'TagName'];
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.webUrl)
    );
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