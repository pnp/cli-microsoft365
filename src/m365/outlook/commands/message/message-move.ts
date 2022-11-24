import * as os from 'os';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { Outlook } from '../../Outlook';

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

      const requestOptions: any = {
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

  private getFolderId(folderId: string | undefined, folderName: string | undefined): Promise<string> {
    if (folderId) {
      return Promise.resolve(folderId);
    }

    if (Outlook.wellKnownFolderNames.indexOf(folderName as string) > -1) {
      return Promise.resolve(folderName as string);
    }

    return new Promise<string>((resolve: (folderId: string) => void, reject: (error: any) => void): void => {
      const requestOptions: any = {
        url: `${this.resource}/v1.0/me/mailFolders?$filter=displayName eq '${formatting.encodeQueryParameter(folderName as string)}'&$select=id`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      request
        .get<{ value: { id: string; }[] }>(requestOptions)
        .then((response: { value: { id: string; }[] }): void => {
          if (response.value.length === 1) {
            return resolve(response.value[0].id);
          }

          if (response.value.length === 0) {
            return reject(`Folder with name '${folderName as string}' not found`);
          }

          if (response.value.length > 1) {
            return reject(`Multiple folders with name '${folderName as string}' found. Please disambiguate:${os.EOL}${response.value.map(f => `- ${f.id}`).join(os.EOL)}`);
          }
        }, err => reject(err));
    });
  }
}

module.exports = new OutlookMessageMoveCommand();
