import * as chalk from 'chalk';
import request from '../../../../request';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import Utils from '../../../../Utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import {FileSharingPrincipalType } from './FileSharingPrincipalType';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  id?: string;
  url?: string;
}

class SpoFileSharinginfoGetCommand extends SpoCommand {

  public get name(): string {
    return commands.FILE_SHARINGINFO_GET;
  }

  public get description(): string {
    return 'Generates the sharing information report for specified file';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = (!(!args.options.id)).toString();
    telemetryProps.url = (!(!args.options.url)).toString();
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.log(`Retrieving sharing information report for the file...`);
    }

    this.getneededFileInformation(args, logger)
      .then((fileInformation: { fileItemId: number; libraryName: string; }): Promise<string> => {
        if (this.verbose) {
          logger.log(`Retrieving Sharing information report for the file with item Id  ${fileInformation.fileItemId}`);
        }

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/lists/getbytitle('${fileInformation.libraryName}')/items(${fileInformation.fileItemId})/GetSharingInformation?$select=permissionsInformation&$Expand=permissionsInformation`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };
        return request.post(requestOptions)
      }).then((res: any): void => {
        if (args.options.output === 'json') {
          logger.log(res);
        }
        else {
          let fileSharingInfoCollection: any[] = [];

          res.permissionsInformation.links.map((links: any) => {
            links.linkDetails.Invitations.map((linkInvites: any) => {
              const fileSharingInfo: any = {
                SharedWith: linkInvites.invitee.name,
                IsActive: linkInvites.invitee.isActive,
                IsExternal: linkInvites.invitee.isExternal,
                PrincipalType: FileSharingPrincipalType[parseInt(linkInvites.invitee.principalType)]
              };
              fileSharingInfoCollection.push(fileSharingInfo);
            })
          })
          res.permissionsInformation.principals.map((principals: any) => {
            const FileSharingInfo: any = {
              SharedWith: principals.principal.name,
              IsActive: principals.principal.isActive,
              IsExternal: principals.principal.isExternal,
              PrincipalType: FileSharingPrincipalType[parseInt(principals.principal.principalType)]
            };
            fileSharingInfoCollection.push(FileSharingInfo);
          });

          logger.log(fileSharingInfoCollection.map((r: any) => {
            return {
              SharedWith: r.SharedWith,
              IsActive: r.IsActive,
              IsExternal: r.IsExternal,
              PrincipalType: r.PrincipalType,
            }
          }));
        }
        if (this.verbose) {
          logger.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getneededFileInformation(args: CommandArgs, logger: Logger): Promise<{ fileItemId: number; libraryName: string; }> {
    let requestUrl: string = '';
    if (args.options.id) {
      requestUrl = `${args.options.webUrl}/_api/web/GetFileById('${escape(args.options.id as string)}')/?$select=ListItemAllFields/Id,ListItemAllFields/ParentList/Title&$expand=ListItemAllFields/ParentList`;
    }
    else {
      requestUrl = `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(decodedUrl='${encodeURIComponent(args.options.url as string)}')?$select=ListItemAllFields/Id,ListItemAllFields/ParentList/Title&$expand=ListItemAllFields/ParentList`;
    }

    const requestOptions: any = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };
    return request.get<string>(requestOptions)
      .then((res: string): Promise<{ fileItemId: number; libraryName: string; }> => {
        const objResult = JSON.parse(JSON.stringify(res));
        return Promise.resolve({ fileItemId: parseInt(objResult.ListItemAllFields.Id), libraryName: objResult.ListItemAllFields.ParentList.Title });
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-w, --webUrl <webUrl>',
        description: 'The URL of the site where the file is located'
      },
      {
        option: '-i, --id [id]',
        description: 'The UniqueId (Item Id) of the file to retrieve. Specify either url or id but not both'
      },
      {
        option: '-u, --url [url]',
        description: 'The server-relative URL of the file to retrieve. Specify either url or id but not both'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
    if (isValidSharePointUrl !== true) {
      return isValidSharePointUrl;
    }

    if (args.options.id) {
      if (!Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }
    }

    if (args.options.id && args.options.url) {
      return 'Specify id or url, but not both';
    }

    if (!args.options.id && !args.options.url) {
      return 'Specify id or url, one is required';
    }

    return true;
  }
}

module.exports = new SpoFileSharinginfoGetCommand();