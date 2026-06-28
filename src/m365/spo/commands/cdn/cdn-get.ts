import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import config from '../../../../config.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { ClientSvcResponse, ClientSvcResponseContents, spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  type: z.enum(['Public', 'Private']).optional().alias('t')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoCdnGetCommand extends SpoCommand {
  public get name(): string {
    return commands.CDN_GET;
  }

  public get description(): string {
    return 'View current status of the specified Microsoft 365 CDN';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const cdnTypeString: string = args.options.type || 'Public';
    const cdnType: number = cdnTypeString === 'Private' ? 1 : 0;

    try {
      const tenantId = await spo.getTenantId(logger, this.debug);
      const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
      const reqDigest = await spo.getRequestDigest(spoAdminUrl);

      if (this.verbose) {
        await logger.logToStderr(`Retrieving status of ${(cdnType === 1 ? 'Private' : 'Public')} CDN...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': reqDigest.FormDigestValue
        },
        data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="GetTenantCdnEnabled" Id="12" ObjectPathId="8"><Parameters><Parameter Type="Enum">${cdnType}</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="8" Name="${tenantId}" /></ObjectPaths></Request>`
      };

      const res = await request.post<string>(requestOptions);
      const json: ClientSvcResponse = JSON.parse(res);
      const response: ClientSvcResponseContents = json[0];
      if (response.ErrorInfo) {
        throw response.ErrorInfo.ErrorMessage;
      }

      const result: boolean = json[json.length - 1];
      if (this.verbose) {
        await logger.logToStderr(`${cdnType === 0 ? 'Public' : 'Private'} CDN at ${spoAdminUrl} is ${result === true ? 'enabled' : 'disabled'}`);
      }

      await logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }
}

export default new SpoCdnGetCommand();