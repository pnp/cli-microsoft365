import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { ClientSvcResponse, ClientSvcResponseContents, ContextInfo, IdentityResponse, spo, urlUtil, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    let formDigestValue: string = '';

    spo
      .getRequestDigest(args.options.webUrl)
      .then((contextResponse: ContextInfo): Promise<IdentityResponse> => {
        formDigestValue = contextResponse.FormDigestValue;

        return spo.getCurrentWebIdentity(args.options.webUrl, formDigestValue);
      })
      .then((webIdentityResp: IdentityResponse): Promise<IdentityResponse> => {
        return spo.getFolderIdentity(webIdentityResp.objectIdentity, args.options.webUrl, args.options.folderUrl, formDigestValue);
      })
      .then((folderObjectIdentity: IdentityResponse): Promise<void> => {
        if (this.verbose) {
          logger.logToStderr(`Renaming folder ${args.options.folderUrl} to ${args.options.name}`);
        }

        const serverRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.folderUrl);
        const serverRelativeUrlWithoutOldFolder: string = serverRelativeUrl.substring(0, serverRelativeUrl.lastIndexOf('/'));
        const renamedServerRelativeUrl: string = `${serverRelativeUrlWithoutOldFolder}/${args.options.name}`;

        const requestOptions: any = {
          url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': formDigestValue
          },
          data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="MoveTo" Id="32" ObjectPathId="26"><Parameters><Parameter Type="String">${renamedServerRelativeUrl}</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="26" Name="${folderObjectIdentity.objectIdentity}" /></ObjectPaths></Request>`
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
      .then(_ => cb(), (err: any): void => this.handleRejectedPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-f, --folderUrl <folderUrl>'
      },
      {
        option: '-n, --name <name>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    return validation.isValidSharePointUrl(args.options.webUrl);
  }
}

module.exports = new SpoFolderRenameCommand();