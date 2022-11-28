import { Logger } from "../../../../cli/Logger";
import GlobalOptions from "../../../../GlobalOptions";
import { odata } from "../../../../utils/odata";
import { validation } from "../../../../utils/validation";
import SpoCommand from "../../../base/SpoCommand";
import commands from "../../commands";
import { formatting } from '../../../../utils/formatting';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
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
      logger.logToStderr(`Retrieving all retention labels that are available on ${args.options.webUrl}...`);
    }

    const requestUrl: string = `${args.options.webUrl}/_api/SP.CompliancePolicy.SPPolicyStoreProxy.GetAvailableTagsForSite(siteUrl=@a1)?@a1='${formatting.encodeQueryParameter(args.options.webUrl)}'`;

    try {
      const response = await odata.getAllItems(requestUrl);
      logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoWebRetentionLabelListCommand();