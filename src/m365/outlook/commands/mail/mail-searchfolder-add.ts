import { MailSearchFolder } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { validation } from '../../../../utils/validation.js';
import { accessToken } from '../../../../utils/accessToken.js';
import auth from '../../../../Auth.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  userId: z.uuid().optional().alias('i'),
  userName: z.string()
    .refine(userName => validation.isValidUserPrincipalName(userName), {
      error: e => `'${e.input}' is not a valid UPN.`
    }).optional().alias('n'),
  folderName: z.string(),
  messageFilter: z.string(),
  sourceFoldersIds: z.string().transform((value) => value.split(',')).pipe(z.string().array()),
  includeNestedFolders: z.boolean().optional()
});

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

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => !(options.userId && options.userName), {
        error: 'Specify either userId or userName, but not both'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken);

      let requestUrl = `${this.resource}/v1.0/me/mailFolders/searchFolders/childFolders`;

      if (isAppOnlyAccessToken) {
        if (!args.options.userId && !args.options.userName) {
          throw 'When running with application permissions either userId or userName is required';
        }

        const userIdentifier = args.options.userId ?? args.options.userName;

        requestUrl = `${this.resource}/v1.0/users('${userIdentifier}')/mailFolders/searchFolders/childFolders`;

        if (args.options.verbose) {
          await logger.logToStderr(`Creating a mail search folder in the mailbox of the user ${userIdentifier}...`);
        }
      }
      else {
        if (args.options.userId || args.options.userName) {
          throw 'You can create mail search folder for other users only if CLI is authenticated in app-only mode';
        }
      }

      const requestOptions: CliRequestOptions = {
        url: requestUrl,
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