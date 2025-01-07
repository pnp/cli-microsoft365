import { MailSearchFolder } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraUser } from '../../../../utils/entraUser.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { zod } from '../../../../utils/zod.js';
import { validation } from '../../../../utils/validation.js';

const options = globalOptionsZod
  .extend({
    userId: zod.alias('i', z.string()
      .refine(userId => validation.isValidGuid(userId), userId => ({
        message: `'${userId}' is not a valid GUID.`
      })).optional()),
    userName: zod.alias('n', z.string()
      .refine(userName => validation.isValidUserPrincipalName(userName), userName => ({
        message: `'${userName}' is not a valid UPN.`
      })).optional()),
    folderName: z.string(),
    messageFilter: z.string(),
    sourceFoldersIds: z.string().transform((value) => value.split(',')).pipe(z.string().array()),
    includeNestedFolders: z.boolean().optional()
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class OutlookMailSearchFolderAddCommand extends GraphCommand {
  public get name(): string {
    return commands.MAIL_SEARCHFOLDER_ADD;
  }

  public get description(): string {
    return `Creates a new mail search folder in the user's mailbox`;
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => !options.userId !== !options.userName, {
        message: 'Specify either userId or userName, but not both'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let userId = args.options.userId;

      if (args.options.userName) {
        userId = await entraUser.getUserIdByUpn(args.options.userName);
      }

      if (args.options.verbose) {
        await logger.logToStderr(`Creating a mail search folder in the mailbox of the user ${userId}...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/users/${userId}/mailFolders/searchFolders/childFolders`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json'
        },
        responseType: 'json',
        data: {
          '@odata.type': '#microsoft.graph.mailSearchFolder',
          displayName: args.options.folderName,
          includeNestedFolders: args.options.includeNestedFolders,
          filterQuery: args.options.messageFilter,
          sourceFolderIds: args.options.sourceFoldersIds
        }
      };

      const result = await request.post<MailSearchFolder>(requestOptions);
      await logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new OutlookMailSearchFolderAddCommand();