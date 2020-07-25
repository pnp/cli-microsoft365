import commands from '../../commands';
import * as os from 'os';
import * as chalk from 'chalk';
import request from '../../../../request';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import { Outlook } from '../../Outlook';
import GraphCommand from '../../../base/GraphCommand';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  messageId: string;
  sourceFolderId?: string;
  sourceFolderName?: string;
  targetFolderId?: string;
  targetFolderName?: string;
}

class OutlookMessageMoveCommand extends GraphCommand {
  public get name(): string {
    return `${commands.OUTLOOK_MESSAGE_MOVE}`;
  }

  public get description(): string {
    return 'Moves message to the specified folder';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.sourceFolderId = typeof args.options.sourceFolderId !== 'undefined';
    telemetryProps.sourceFolderName = typeof args.options.sourceFolderName !== 'undefined';
    telemetryProps.targetFolderId = typeof args.options.targetFolderId !== 'undefined';
    telemetryProps.targetFolderName = typeof args.options.targetFolderName !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    let sourceFolder: string;
    let targetFolder: string;

    this
      .getFolderId(args.options.sourceFolderId, args.options.sourceFolderName)
      .then((folderId: string): Promise<string> => {
        sourceFolder = folderId;

        return this.getFolderId(args.options.targetFolderId, args.options.targetFolderName);
      })
      .then((folderId: string): Promise<void> => {
        targetFolder = folderId;

        const messageUrl: string = `mailFolders/${sourceFolder}/messages/${args.options.messageId}`;

        const requestOptions: any = {
          url: `${this.resource}/v1.0/me/${messageUrl}/move`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          body: {
            destinationId: targetFolder
          },
          json: true
        };

        return request.post(requestOptions);
      })
      .then((): void => {
        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
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
        url: `${this.resource}/v1.0/me/mailFolders?$filter=displayName eq '${encodeURIComponent(folderName as string)}'&$select=id`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        json: true
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

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--messageId <messageId>',
        description: 'ID of the message to move'
      },
      {
        option: '--sourceFolderName [sourceFolderName]',
        description: 'Name of the folder to move the message from',
        autocomplete: Outlook.wellKnownFolderNames
      },
      {
        option: '--sourceFolderId [sourceFolderId]',
        description: 'ID of the folder to move the message from',
        autocomplete: Outlook.wellKnownFolderNames
      },
      {
        option: '--targetFolderName [targetFolderName]',
        description: 'Name of the folder to move the message to',
        autocomplete: Outlook.wellKnownFolderNames
      },
      {
        option: '--targetFolderId [targetFolderId]',
        description: 'ID of the folder to move the message to',
        autocomplete: Outlook.wellKnownFolderNames
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.sourceFolderId &&
        !args.options.sourceFolderName) {
        return 'Specify sourceFolderId or sourceFolderName';
      }

      if (args.options.sourceFolderId &&
        args.options.sourceFolderName) {
        return 'Specify either sourceFolderId or sourceFolderName but not both';
      }

      if (!args.options.targetFolderId &&
        !args.options.targetFolderName) {
        return 'Specify targetFolderId or targetFolderName';
      }

      if (args.options.targetFolderId &&
        args.options.targetFolderName) {
        return 'Specify either targetFolderId or targetFolderName but not both';
      }

      return true;
    };
  }
}

module.exports = new OutlookMessageMoveCommand();
