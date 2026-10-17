import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { FileSharingPrincipalType } from './FileSharingPrincipalType.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  fileId: z.uuid().optional().alias('i'),
  fileUrl: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

interface SharingPrincipal {
  isActive: boolean;
  isExternal: boolean;
  name: string;
  principalType: string;
}

interface SharingInformation {
  permissionsInformation: {
    links: {
      linkDetails: {
        Invitations: {
          invitee: SharingPrincipal;
        }[];
      };
    }[];
    principals: {
      principal: SharingPrincipal;
    }[];
  };
}

interface FileSharingInformation {
  IsActive: boolean;
  IsExternal: boolean;
  PrincipalType: string;
  SharedWith: string;
}

class SpoFileSharingInfoGetCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_SHARINGINFO_GET;
  }

  public get description(): string {
    return 'Generates a sharing information report for the specified file';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.fileId, opts.fileUrl].filter(o => o !== undefined).length === 1, {
        error: `Specify 'fileId' or 'fileUrl', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['fileId', 'fileUrl']
        }
      });
  }

  protected getExcludedOptionsWithUrls(): string[] | undefined {
    return ['fileUrl'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving sharing information report for the file ${args.options.fileId || args.options.fileUrl}`);
    }

    try {
      const fileInformation = await this.getNeededFileInformation(args);
      if (this.verbose) {
        await logger.logToStderr(`Retrieving sharing information report for the file with item Id  ${fileInformation.fileItemId}`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/web/lists/getbytitle('${formatting.encodeQueryParameter(fileInformation.libraryName)}')/items(${fileInformation.fileItemId})/GetSharingInformation?$select=permissionsInformation&$Expand=permissionsInformation`,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };
      const res = await request.post<SharingInformation>(requestOptions);

      // typically, we don't do this, but in this case, we need to due to
      // the complexity of the retrieved object and the fact that we can't
      // use the generic way of simplifying the output
      if (!cli.shouldTrimOutput(args.options.output)) {
        await logger.log(res);
      }
      else {
        const fileSharingInfoCollection: FileSharingInformation[] = [];
        res.permissionsInformation.links.forEach(link => {
          link.linkDetails.Invitations.forEach(linkInvite => {
            fileSharingInfoCollection.push({
              SharedWith: linkInvite.invitee.name,
              IsActive: linkInvite.invitee.isActive,
              IsExternal: linkInvite.invitee.isExternal,
              PrincipalType: FileSharingPrincipalType[parseInt(linkInvite.invitee.principalType)]
            });
          });
        });
        res.permissionsInformation.principals.forEach(principal => {
          fileSharingInfoCollection.push({
            SharedWith: principal.principal.name,
            IsActive: principal.principal.isActive,
            IsExternal: principal.principal.isExternal,
            PrincipalType: FileSharingPrincipalType[parseInt(principal.principal.principalType)]
          });
        });

        await logger.log(fileSharingInfoCollection);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getNeededFileInformation(args: CommandArgs): Promise<{ fileItemId: number; libraryName: string; }> {
    let requestUrl: string;

    if (args.options.fileId) {
      requestUrl = `${args.options.webUrl}/_api/web/GetFileById('${args.options.fileId}')/?$select=ListItemAllFields/Id,ListItemAllFields/ParentList/Title&$expand=ListItemAllFields/ParentList`;
    }
    else {
      const serverRelPath = urlUtil.getServerRelativePath(args.options.webUrl, args.options.fileUrl!);
      requestUrl = `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(serverRelPath)}')?$select=ListItemAllFields/Id,ListItemAllFields/ParentList/Title&$expand=ListItemAllFields/ParentList`;
    }

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const res = await request.get<{ ListItemAllFields: { Id: string; ParentList: { Title: string }; } }>(requestOptions);
    return {
      fileItemId: parseInt(res.ListItemAllFields.Id),
      libraryName: res.ListItemAllFields.ParentList.Title
    };
  }
}

export default new SpoFileSharingInfoGetCommand();