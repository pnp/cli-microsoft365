import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  siteUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  fileUrl: z.string()
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoReportSiteVersionStorageUsageStatusGetCommand extends SpoCommand {
  public get name(): string {
    return commands.REPORT_SITEVERSIONSTORAGEUSAGE_STATUS_GET;
  }

  public get description(): string {
    return 'Retrieves the status of a version storage usage report';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving status of version storage usage report for site '${args.options.siteUrl}'...`);
    }

    try {
      const requestOptions: CliRequestOptions = {
        url: `${args.options.siteUrl}/_api/site/GetProgressForFileVersionExpirationReport(reportFileUrl='${formatting.encodeQueryParameter(args.options.fileUrl)}')`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const response = await request.post<{ value: string }>(requestOptions);
      await logger.log(JSON.parse(response.value));
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoReportSiteVersionStorageUsageStatusGetCommand();