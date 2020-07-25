import commands from '../../commands';
import * as os from 'os';
import * as chalk from 'chalk';
import request from '../../../../request';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import { Message } from '../../Message';
import { Outlook } from '../../Outlook';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  folderId?: string;
  folderName?: string;
}

class OutlookMessageListCommand extends GraphItemsListCommand<Message> {
  public get name(): string {
    return `${commands.OUTLOOK_MESSAGE_LIST}`;
  }

  public get description(): string {
    return 'Gets all mail messages from the specified folder';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.folderId = typeof args.options.folderId !== 'undefined';
    telemetryProps.folderName = typeof args.options.folderName !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    this
      .getFolderId(args)
      .then((folderId: string): Promise<void> => {
        const url: string = folderId ? `me/mailFolders/${folderId}/messages` : 'me/messages';

        return this.getAllItems(`${this.resource}/v1.0/${url}?$top=50`, cmd, true);
      })
      .then((): void => {
        if (args.options.output === 'json') {
          cmd.log(this.items);
        }
        else {
          cmd.log(this.items.map(i => {
            return {
              subject: i.subject,
              receivedDateTime: i.receivedDateTime
            };
          }));
        }

        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
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
        url: `${this.resource}/v1.0/me/mailFolders?$filter=displayName eq '${encodeURIComponent(args.options.folderName as string)}'&$select=id`,
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
            return reject(`Folder with name '${args.options.folderName as string}' not found`);
          }

          if (response.value.length > 1) {
            return reject(`Multiple folders with name '${args.options.folderName as string}' found. Please disambiguate:${os.EOL}${response.value.map(f => `- ${f.id}`).join(os.EOL)}`);
          }
        }, err => reject(err));
    });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--folderName [folderName]',
        description: 'Name of the folder from which to list messages',
        autocomplete: Outlook.wellKnownFolderNames
      },
      {
        option: '--folderId [folderId]',
        description: 'ID of the folder from which to list messages',
        autocomplete: Outlook.wellKnownFolderNames
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.folderId &&
        !args.options.folderName) {
        return 'Specify folderId or folderName';
      }

      if (args.options.folderId &&
        args.options.folderName) {
        return 'Specify either folderId or folderName but not both';
      }

      return true;
    };
  }
}

module.exports = new OutlookMessageListCommand();
