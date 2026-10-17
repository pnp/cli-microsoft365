import { Group } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { FileProperties } from './FileProperties.js';
import { entraGroup } from '../../../../utils/entraGroup.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  fileUrl: z.string().optional(),
  fileId: z.uuid().optional().alias('i'),
  principalId: z.number().optional(),
  upn: z.string().optional(),
  groupName: z.string().optional(),
  entraGroupId: z.string().refine(id => validation.isValidGuid(id), {
    error: e => `'${e.input}' is not a valid GUID for option entraGroupId.`
  }).optional(),
  entraGroupName: z.string().optional(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFileRoleAssignmentRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_ROLEASSIGNMENT_REMOVE;
  }

  public get description(): string {
    return 'Removes a role assignment from a file.';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.fileUrl, opts.fileId].filter(o => o !== undefined).length === 1, {
        error: `Specify 'fileUrl' or 'fileId', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['fileUrl', 'fileId']
        }
      })
      .refine(opts => [opts.upn, opts.groupName, opts.principalId, opts.entraGroupId, opts.entraGroupName].filter(o => o !== undefined).length === 1, {
        error: `Specify either 'upn', 'groupName', 'principalId', 'entraGroupId', or 'entraGroupName'.`,
        params: {
          customCode: 'optionSet',
          options: ['upn', 'groupName', 'principalId', 'entraGroupId', 'entraGroupName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeRoleAssignment = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Removing role assignment for ${args.options.groupName || args.options.upn} from file ${args.options.fileUrl || args.options.fileId}`);
      }

      try {
        const fileURL: string = await this.getFileURL(args, logger);

        let principalId: number;
        if (args.options.groupName) {
          principalId = await this.getGroupPrincipalId(args.options, logger);
        }
        else if (args.options.upn) {
          principalId = await this.getUserPrincipalId(args.options, logger);
        }
        else if (args.options.entraGroupId || args.options.entraGroupName) {
          if (this.verbose) {
            await logger.logToStderr('Retrieving group information...');
          }

          let group: Group;
          if (args.options.entraGroupId) {
            group = await entraGroup.getGroupById(args.options.entraGroupId);
          }
          else {
            group = await entraGroup.getGroupByDisplayName(args.options.entraGroupName!);
          }

          const entraSiteUser = await spo.ensureEntraGroup(args.options.webUrl, group);

          principalId = entraSiteUser.Id;
        }
        else {
          principalId = args.options.principalId!;
        }

        const serverRelativePath: string = urlUtil.getServerRelativePath(args.options.webUrl, fileURL);
        const requestOptions: CliRequestOptions = {
          url: `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')/ListItemAllFields/roleassignments/removeroleassignment(principalid='${principalId}')`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        await request.post(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeRoleAssignment();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove role assignment from file ${args.options.fileUrl || args.options.fileId} from site ${args.options.webUrl}?` });

      if (result) {
        await removeRoleAssignment();
      }
    }
  }

  private async getFileURL(args: CommandArgs, logger: Logger): Promise<string> {
    if (args.options.fileUrl) {
      return urlUtil.getServerRelativePath(args.options.webUrl, args.options.fileUrl);
    }

    const file: FileProperties = await spo.getFileById(args.options.webUrl, args.options.fileId!, logger, this.verbose);
    return file.ServerRelativeUrl;
  }

  private async getUserPrincipalId(options: Options, logger: Logger): Promise<number> {
    const user = await spo.getUserByEmail(options.webUrl, options.upn!, logger, this.verbose);
    return user.Id;
  }

  private async getGroupPrincipalId(options: Options, logger: Logger): Promise<number> {
    const group = await spo.getGroupByName(options.webUrl, options.groupName!, logger, this.verbose);
    return group.Id;
  }
}

export default new SpoFileRoleAssignmentRemoveCommand();