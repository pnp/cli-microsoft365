import * as request from 'request-promise-native';
import Utils from '../../../../Utils';
import * as url from 'url';

/**
 * Folder methods that are shared among multiple commands.
 */
export class FolderExtensions {

  public constructor(private cmd: CommandInstance, private debug: boolean) {
  }

  public ensureFolder(webFullUrl: string, folderToEnsure: string, siteAccessToken: string): Promise<void> {

    const webUrl = url.parse(webFullUrl);
    if(!webUrl.protocol || !webUrl.hostname) {
      return Promise.reject('webFullUrl is not a valid URL');
    }

    if(!folderToEnsure){
      return Promise.reject('folderToEnsure cannot be empty');
    }

    if(!siteAccessToken){
      return Promise.reject('siteAccessToken cannot be empty');
    }

    // remove the end '/' in the folder path
    if (folderToEnsure[folderToEnsure.length - 1] === '/') {
      folderToEnsure = folderToEnsure.substring(0, folderToEnsure.length - 1);
    }

    const tenantUrl: string = `${url.parse(webFullUrl).protocol}//${url.parse(webFullUrl).hostname}`;
    const webRelativePath: string = webFullUrl.replace(tenantUrl, '');

    // remove the web relative path from the folder path
    folderToEnsure = folderToEnsure.replace(webRelativePath, '')

    // remove the leading '/' so we are left with e.g. Shared%20Documents/22/54/55
    if (folderToEnsure[0] === '/') {
      folderToEnsure = folderToEnsure.substring(1);
    }

    if (this.debug) {
      this.cmd.log(`folderToEnsure`);
      this.cmd.log(folderToEnsure);
      this.cmd.log('');
    }

    // build array of folders e.g. ["Shared%20Documents","22","54","55"]
    let folders: string[] = folderToEnsure.split('/');
    let nextFolder: string = '';
    let folderIndex: number = 0;

    // recursive function
    const checkOrAddFolder = (resolve: () => void, reject: (error: any) => void): void => {

      if (folderIndex === folders.length) {

        if (this.debug) {
          this.cmd.log(`All sub-folders exist`);
        }

        return resolve();
      }

      // append the next sub-folder to the folder path and check if it exists
      nextFolder += `/${folders[folderIndex]}`;
      const folderServerRelativeUrl = `${webRelativePath}${nextFolder}`;

      const requestOptions: any = {
        url: `${webFullUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderServerRelativeUrl)}')`,
        headers: Utils.getRequestHeaders({
          authorization: `Bearer ${siteAccessToken}`,
          'accept': 'application/json;odata=nometadata'
        })
      };

      if (this.debug) {
        this.cmd.log(`Check if ${nextFolder} exists`);
        this.cmd.log(requestOptions);
        this.cmd.log('');
      }

      request.get(requestOptions)
        .then((res: any) => {

          if (this.debug) {
            this.cmd.log(`${nextFolder} exists. Moving to the next one`);
            this.cmd.log(res);
            this.cmd.log('');
          }

          folderIndex++;
          checkOrAddFolder(resolve, reject);
        })
        .catch((err: any) => {
          
          const requestOptions: any = {
            url: `${webFullUrl}/_api/web/folders`,
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${siteAccessToken}`,
              'accept': 'application/json;odata=nometadata',
            }),
            body: {
              'ServerRelativeUrl': folderServerRelativeUrl
            },
            json: true
          };

          if (this.debug) {
            this.cmd.log(`Add folder ${folderServerRelativeUrl}`);
            this.cmd.log(requestOptions);
            this.cmd.log('');
          }

          return request.post(requestOptions)
            .then((res: any) => {

              if (this.debug) {
                this.cmd.log(`Folder ${folderServerRelativeUrl} added`);
                this.cmd.log(JSON.stringify(res));
                this.cmd.log('');
              }

              folderIndex++;
              checkOrAddFolder(resolve, reject);
            })
            .catch((err: any) => {

              if (this.debug) {
                this.cmd.log(`Could not create sub-folder ${folderServerRelativeUrl}`);
              }

              reject(err); 
            });
        });
    }
    return new Promise<void>(checkOrAddFolder);
  }
}