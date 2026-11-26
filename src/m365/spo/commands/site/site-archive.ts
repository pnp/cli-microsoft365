import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import config from '../../../../config.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import { ClientSvcResponse, ClientSvcResponseContents, spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  url: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoSiteArchiveCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_ARCHIVE;
  }

  public get description(): string {
    return 'Archives a site collection';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const archiveSite = async (): Promise<void> => {
      try {
        if (this.verbose) {
          await logger.logToStderr(`Archiving site ${args.options.url}...`);
        }

        const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.verbose);
        const reqDigest = await spo.getRequestDigest(spoAdminUrl);

        const requestOptions: CliRequestOptions = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': reqDigest.FormDigestValue
          },
          data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009">
                  <Actions>
                    <ObjectPath Id="2" ObjectPathId="1" />
                    <ObjectPath Id="4" ObjectPathId="3" />
                    <Query Id="5" ObjectPathId="3">
                      <Query SelectAllProperties="true">
                        <Properties />
                      </Query>
                    </Query>
                  </Actions>
                  <ObjectPaths>
                    <Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" />
                    <Method Id="3" ParentId="1" Name="ArchiveSiteByUrl">
                      <Parameters>
                        <Parameter Type="String">${args.options.url}</Parameter>
                      </Parameters>
                    </Method>
                  </ObjectPaths>
                  </Request>`
        };

        const res = await request.post<string>(requestOptions);

        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          throw response.ErrorInfo.ErrorMessage;
        }
      }
      catch (err: any) {
        this.handleRejectedPromise(err);
      }
    };

    if (args.options.force) {
      await archiveSite();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to archive site '${args.options.url}'?` });

      if (result) {
        await archiveSite();
      }
    }
  }
}

export default new SpoSiteArchiveCommand();