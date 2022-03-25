import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { FileSharingPrincipalType } from './FileSharingPrincipalType';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  id?: string;
  url?: string;
}

interface SharingPrincipal {
  isActive: boolean;
  isExternal: boolean;
  name: string;
  principalType: string;
}

interface SharingInformation {
  permissionsInformation: {
    links: {
      linkDetails: {
        Invitations: {
          invitee: SharingPrincipal;
        }[];
      };
    }[];
    principals: {
      principal: SharingPrincipal;
    }[];
  };
}

interface FileSharingInformation {
  IsActive: boolean;
  IsExternal: boolean;
  PrincipalType: string;
  SharedWith: string;
}

class SpoFileSharinginfoGetCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_SHARINGINFO_GET;
  }

  public get description(): string {
    return 'Generates a sharing information report for the specified file';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = (!(!args.options.id)).toString();
    telemetryProps.url = (!(!args.options.url)).toString();
    return telemetryProps;
  }

  protected getExcludedOptionsWithUrls(): string[] | undefined {
    return ['url'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving sharing information report for the file...`);
    }

    this
      .getNeededFileInformation(args)
      .then((fileInformation: { fileItemId: number; libraryName: string; }): Promise<SharingInformation> => {
        if (this.verbose) {
          logger.logToStderr(`Retrieving sharing information report for the file with item Id  ${fileInformation.fileItemId}`);
        }

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/lists/getbytitle('${fileInformation.libraryName}')/items(${fileInformation.fileItemId})/GetSharingInformation?$select=permissionsInformation&$Expand=permissionsInformation`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };
        return request.post(requestOptions);
      }).then((res: SharingInformation): void => {
        // typically, we don't do this, but in this case, we need to due to
        // the complexity of the retrieved object and the fact that we can't
        // use the generic way of simplifying the output
        if (args.options.output === 'json') {
          logger.log(res);
        }
        else {
          const fileSharingInfoCollection: FileSharingInformation[] = [];
          res.permissionsInformation.links.forEach(link => {
            link.linkDetails.Invitations.forEach(linkInvite => {
              fileSharingInfoCollection.push({
                SharedWith: linkInvite.invitee.name,
                IsActive: linkInvite.invitee.isActive,
                IsExternal: linkInvite.invitee.isExternal,
                PrincipalType: FileSharingPrincipalType[parseInt(linkInvite.invitee.principalType)]
              });
            });
          });
          res.permissionsInformation.principals.forEach(principal => {
            fileSharingInfoCollection.push({
              SharedWith: principal.principal.name,
              IsActive: principal.principal.isActive,
              IsExternal: principal.principal.isExternal,
              PrincipalType: FileSharingPrincipalType[parseInt(principal.principal.principalType)]
            });
          });

          logger.log(fileSharingInfoCollection);
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getNeededFileInformation(args: CommandArgs): Promise<{ fileItemId: number; libraryName: string; }> {
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

    return request.get<{ ListItemAllFields: { Id: string; ParentList: { Title: string }; } }>(requestOptions)
      .then((res: { ListItemAllFields: { Id: string; ParentList: { Title: string }; } }): Promise<{ fileItemId: number; libraryName: string; }> => Promise.resolve({
        fileItemId: parseInt(res.ListItemAllFields.Id),
        libraryName: res.ListItemAllFields.ParentList.Title
      }));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-w, --webUrl <webUrl>'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '-u, --url [url]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
    if (isValidSharePointUrl !== true) {
      return isValidSharePointUrl;
    }

    if (args.options.id) {
      if (!validation.isValidGuid(args.options.id)) {
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