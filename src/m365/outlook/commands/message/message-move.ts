import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import commands from '../../commands.js';
import { Outlook } from '../../Outlook.js';
import { cli } from '../../../../cli/cli.js';
import DelegatedGraphCommand from '../../../base/GraphDelegatedCommand.js';
import { z } from 'zod';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string(),
  sourceFolderId: z.string().optional(),
  sourceFolderName: z.string().optional(),
  targetFolderId: z.string().optional(),
  targetFolderName: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class OutlookMessageMoveCommand extends DelegatedGraphCommand {
  public get name(): string {
    return commands.MESSAGE_MOVE;
  }

  public get description(): string {
    return 'Moves message to the specified folder';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => options.sourceFolderId || options.sourceFolderName, {
        error: 'Specify either sourceFolderId or sourceFolderName',
        params: {
          customCode: 'optionSet',
          options: ['sourceFolderId', 'sourceFolderName']
        }
      })
      .refine(options => !(options.sourceFolderId && options.sourceFolderName), {
        error: 'Specify either sourceFolderId or sourceFolderName, but not both',
        params: {
          customCode: 'optionSet',
          options: ['sourceFolderId', 'sourceFolderName']
        }
      })
      .refine(options => options.targetFolderId || options.targetFolderName, {
        error: 'Specify either targetFolderId or targetFolderName',
        params: {
          customCode: 'optionSet',
          options: ['targetFolderId', 'targetFolderName']
        }
      })
      .refine(options => !(options.targetFolderId && options.targetFolderName), {
        error: 'Specify either targetFolderId or targetFolderName, but not both',
        params: {
          customCode: 'optionSet',
          options: ['targetFolderId', 'targetFolderName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let sourceFolder: string;
    let targetFolder: string;

    try {
      sourceFolder = await this.getFolderId(args.options.sourceFolderId, args.options.sourceFolderName);
      targetFolder = await this.getFolderId(args.options.targetFolderId, args.options.targetFolderName);

      const messageUrl: string = `mailFolders/${sourceFolder}/messages/${args.options.id}`;

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/me/${messageUrl}/move`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        data: {
          destinationId: targetFolder
        },
        responseType: 'json'
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getFolderId(folderId: string | undefined, folderName: string | undefined): Promise<string> {
    if (folderId) {
      return folderId;
    }

    if (Outlook.wellKnownFolderNames.indexOf(folderName as string) > -1) {
      return folderName as string;
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/me/mailFolders?$filter=displayName eq '${formatting.encodeQueryParameter(folderName as string)}'&$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: { id: string; }[] }>(requestOptions);

    if (response.value.length === 0) {
      throw `Folder with name '${folderName as string}' not found`;
    }

    if (response.value.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', response.value);
      const result = await cli.handleMultipleResultsFound<{ id: string; }>(`Multiple folders with name '${folderName as string}' found.`, resultAsKeyValuePair);
      return result.id;
    }

    return response.value[0].id;
  }
}

export default new OutlookMessageMoveCommand();
