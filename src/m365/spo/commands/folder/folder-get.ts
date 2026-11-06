import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError, globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { ListPrincipalType } from '../list/ListPrincipalType.js';
import { FolderProperties } from './FolderProperties.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  url: z.string().optional(),
  id: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `'${e.input}' is not a valid GUID.`
    })
    .optional()
    .alias('i'),
  withPermissions: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFolderGetCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_GET;
  }

  public get description(): string {
    return 'Gets information about the specified folder';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.url, opts.id].filter(o => o !== undefined).length === 1, {
        error: `Specify either 'url' or 'id', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['url', 'id']
        }
      });
  }

  protected getExcludedOptionsWithUrls(): string[] | undefined {
    return ['url'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving folder from site ${args.options.webUrl}...`);
    }
    let requestUrl: string = `${args.options.webUrl}/_api/web`;
    if (args.options.id) {
      requestUrl += `/GetFolderById('${formatting.encodeQueryParameter(args.options.id)}')`;
    }
    else if (args.options.url) {
      const serverRelativePath: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.url);
      requestUrl += `/GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')`;
    }
    if (args.options.withPermissions) {
      requestUrl += `?$expand=ListItemAllFields/HasUniqueRoleAssignments,ListItemAllFields/RoleAssignments/Member,ListItemAllFields/RoleAssignments/RoleDefinitionBindings`;
    }

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const folder = await request.get<FolderProperties>(requestOptions);
      if (args.options.withPermissions) {
        const listItemAllFields = folder.ListItemAllFields;
        if (!(listItemAllFields ?? false)) {
          throw Error('Please ensure the specified folder URL or folder Id does not refer to a root folder. Use \'spo list get\' with withPermissions instead.');
        }
        listItemAllFields.RoleAssignments.forEach(r => {
          r.Member.PrincipalTypeString = ListPrincipalType[r.Member.PrincipalType];
          r.RoleDefinitionBindings = formatting.setFriendlyPermissions(r.RoleDefinitionBindings);
        });
      }
      await logger.log(folder);
    }
    catch (err: any) {
      if (err.statusCode && err.statusCode === 500) {
        throw new CommandError('Please check the folder URL. Folder might not exist on the specified URL');
      }
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoFolderGetCommand();