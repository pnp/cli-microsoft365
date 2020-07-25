import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import { ClientSvc, IdentityResponse } from '../../ClientSvc';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  folderUrl: string;
  name: string;
}

class SpoFolderRenameCommand extends SpoCommand {

  public get name(): string {
    return commands.FOLDER_RENAME;
  }

  public get description(): string {
    return 'Renames a folder';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const clientSvc: ClientSvc = new ClientSvc(cmd, this.debug);
    let formDigestValue: string = '';

    this
      .getRequestDigest(args.options.webUrl)
      .then((contextResponse: ContextInfo): Promise<IdentityResponse> => {
        formDigestValue = contextResponse.FormDigestValue;

        return clientSvc.getCurrentWebIdentity(args.options.webUrl, formDigestValue);
      })
      .then((webIdentityResp: IdentityResponse): Promise<IdentityResponse> => {
        return clientSvc.getFolderIdentity(webIdentityResp.objectIdentity, args.options.webUrl, args.options.folderUrl, formDigestValue);
      })
      .then((folderObjectIdentity: IdentityResponse): Promise<void> => {
        if (this.verbose) {
          cmd.log(`Renaming folder ${args.options.folderUrl} to ${args.options.name}`);
        }

        const serverRelativeUrl: string = Utils.getServerRelativePath(args.options.webUrl, args.options.folderUrl);
        const serverRelativeUrlWithoutOldFolder: string = serverRelativeUrl.substring(0, serverRelativeUrl.lastIndexOf('/'));
        const renamedServerRelativeUrl: string = `${serverRelativeUrlWithoutOldFolder}/${args.options.name}`;

        const requestOptions: any = {
          url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': formDigestValue
          },
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="MoveTo" Id="32" ObjectPathId="26"><Parameters><Parameter Type="String">${renamedServerRelativeUrl}</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="26" Name="${folderObjectIdentity.objectIdentity}" /></ObjectPaths></Request>`
        };

        return new Promise<void>((resolve: any, reject: any): void => {
          request.post(requestOptions).then((res: any) => {
            const json: ClientSvcResponse = JSON.parse(res);
            const contents: ClientSvcResponseContents = json.find(x => { return x['ErrorInfo']; });
            if (contents && contents.ErrorInfo) {
              return reject(contents.ErrorInfo.ErrorMessage || 'ClientSvc unknown error');
            }

            return resolve();
          }, (err: any): void => { reject(err); });
        });
      })
      .then((): void => {
        if (this.verbose) {
          cmd.log('DONE');
        }

        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'The URL of the site where the folder to be renamed is located'
      },
      {
        option: '-f, --folderUrl <folderUrl>',
        description: 'Site-relative URL of the folder (including the folder)'
      },
      {
        option: '-n, --name <name>',
        description: 'New name for the target folder'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      return SpoCommand.isValidSharePointUrl(args.options.webUrl);
    };
  }
}

module.exports = new SpoFolderRenameCommand();