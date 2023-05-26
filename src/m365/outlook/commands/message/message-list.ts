import * as os from 'os';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { odata } from '../../../../utils/odata';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { Outlook } from '../../Outlook';
import { Message } from '@microsoft/microsoft-graph-types';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  folderId?: string;
  folderName?: string;
}

class OutlookMessageListCommand extends GraphCommand {
  public get name(): string {
    return commands.MESSAGE_LIST;
  }

  public get description(): string {
    return 'Gets all mail messages from the specified folder';
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
        folderId: typeof args.options.folderId !== 'undefined',
        folderName: typeof args.options.folderName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--folderName [folderName]',
        autocomplete: Outlook.wellKnownFolderNames
      },
      {
        option: '--folderId [folderId]',
        autocomplete: Outlook.wellKnownFolderNames
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['folderId', 'folderName'] });
  }

  public defaultProperties(): string[] | undefined {
    return ['subject', 'receivedDateTime'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const folderId = await this.getFolderId(args);

      const url: string = folderId ? `me/mailFolders/${folderId}/messages` : 'me/messages';
      const messages = await odata.getAllItems<Message>(`${this.resource}/v1.0/${url}?$top=50`);

      logger.log(messages);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getFolderId(args: CommandArgs): Promise<string> {
    if (!args.options.folderId && !args.options.folderName) {
      return Promise.resolve('');
    }

    if (args.options.folderId) {
      return Promise.resolve(args.options.folderId);
    }

    if (Outlook.wellKnownFolderNames.indexOf(args.options.folderName as string) > -1) {
      return Promise.resolve(args.options.folderName as string);
    }

    return new Promise<string>((resolve: (folderId: string) => void, reject: (error: any) => void): void => {
      const requestOptions: any = {
        url: `${this.resource}/v1.0/me/mailFolders?$filter=displayName eq '${formatting.encodeQueryParameter(args.options.folderName as string)}'&$select=id`,
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
            return reject(`Folder with name '${args.options.folderName as string}' not found`);
          }

          if (response.value.length > 1) {
            return reject(`Multiple folders with name '${args.options.folderName as string}' found. Please disambiguate:${os.EOL}${response.value.map(f => `- ${f.id}`).join(os.EOL)}`);
          }
        }, err => reject(err));
    });
  }
}

module.exports = new OutlookMessageListCommand();
