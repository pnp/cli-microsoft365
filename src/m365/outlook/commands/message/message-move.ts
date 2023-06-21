import os from 'os';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { Outlook } from '../../Outlook.js';
import { Cli } from '../../../../cli/Cli.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  sourceFolderId?: string;
  sourceFolderName?: string;
  targetFolderId?: string;
  targetFolderName?: string;
}

class OutlookMessageMoveCommand extends GraphCommand {
  public get name(): string {
    return commands.MESSAGE_MOVE;
  }

  public get description(): string {
    return 'Moves message to the specified folder';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        sourceFolderId: typeof args.options.sourceFolderId !== 'undefined',
        sourceFolderName: typeof args.options.sourceFolderName !== 'undefined',
        targetFolderId: typeof args.options.targetFolderId !== 'undefined',
        targetFolderName: typeof args.options.targetFolderName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--id <id>'
      },
      {
        option: '--sourceFolderName [sourceFolderName]',
        autocomplete: Outlook.wellKnownFolderNames
      },
      {
        option: '--sourceFolderId [sourceFolderId]',
        autocomplete: Outlook.wellKnownFolderNames
      },
      {
        option: '--targetFolderName [targetFolderName]',
        autocomplete: Outlook.wellKnownFolderNames
      },
      {
        option: '--targetFolderId [targetFolderId]',
        autocomplete: Outlook.wellKnownFolderNames
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['sourceFolderId', 'sourceFolderName'] },
      { options: ['targetFolderId', 'targetFolderName'] }
    );
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
      response.value[0] = await Cli.handleMultipleResultsFound<{ id: string; }>(`Multiple folders with name '${folderName as string}' found. Choose the correct ID:`, `Multiple folders with name '${folderName as string}' found. Please disambiguate:${os.EOL}${response.value.map(f => `- ${f.id}`).join(os.EOL)}`, resultAsKeyValuePair);
    }

    return response.value[0].id;
  }
}

export default new OutlookMessageMoveCommand();
