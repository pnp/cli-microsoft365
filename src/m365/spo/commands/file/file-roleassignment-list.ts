import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(value => validation.isValidSharePointUrl(value) === true, {
      error: issue => `'${issue.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  fileUrl: z.string().optional(),
  fileId: z.string()
    .refine(value => validation.isValidGuid(value), {
      error: issue => `${issue.input} is not a valid GUID`
    })
    .optional()
    .alias('i')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFileRoleAssignmentListCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_ROLEASSIGNMENT_LIST;
  }

  public get description(): string {
    return 'Lists all role assignments from a specific file.';
  }

  public defaultProperties(): string[] | undefined {
    return ['PrincipalId', 'Member', 'RoleDefinitionBindings'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.fileId, opts.fileUrl].filter(value => value !== undefined).length === 1, {
        error: `Specify either 'fileId' or 'fileUrl', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['fileId', 'fileUrl']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving role assignments for file with ${args.options.fileId || args.options.fileUrl} in site at ${args.options.webUrl}...`);
    }

    try {
      let file;
      if (args.options.fileId) {
        file = await spo.getFileById(args.options.webUrl, args.options.fileId, logger, this.verbose);
      }
      else {
        file = await spo.getFileByUrl(args.options.webUrl, args.options.fileUrl!, logger, this.verbose);
      }

      const fileRoleAssignments: {
        PrincipalId: number;
        Member: { Title: string };
        RoleDefinitionBindings: { Name: string }[];
      }[] = await spo.getFileRoleAssignments(args.options.webUrl, file.ServerRelativeUrl, logger, this.verbose);

      if (args.options.output && args.options.output !== 'json') {
        await logger.log(fileRoleAssignments.map(assignment => ({
          PrincipalId: assignment.PrincipalId,
          Member: assignment.Member.Title,
          RoleDefinitionBindings: assignment.RoleDefinitionBindings.map(role => role.Name).join(', ')
        })));
      }
      else {
        await logger.log(fileRoleAssignments);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoFileRoleAssignmentListCommand();