import * as request from 'request-promise-native';
import Utils from '../../../../Utils';
import * as url from 'url';

/**
 * Folder methods that are shared among multiple commands.
 */
export class FolderExtensions {

  public constructor(private cmd: CommandInstance, private debug: boolean) {
  }

  /**
   * Ensures the folder path exists
   * @param webFullUrl web full url e.g. https://contoso.sharepoint.com/sites/site1
   * @param folderToEnsure web relative or server relative folder path e.g. /Documents/MyFolder or /sites/site1/Documents/MyFolder
   * @param siteAccessToken a valid access token for the site specified in the webFullUrl param
   */
  public ensureFolder(webFullUrl: string, folderToEnsure: string, siteAccessToken: string): Promise<void> {

    const webUrl = url.parse(webFullUrl);
    if (!webUrl.protocol || !webUrl.hostname) {
      return Promise.reject('webFullUrl is not a valid URL');
    }

    if (!folderToEnsure) {
      return Promise.reject('folderToEnsure cannot be empty');
    }

    if (!siteAccessToken) {
      return Promise.reject('siteAccessToken cannot be empty');
    }

    // remove last '/' of webFullUrl if exists
    const webFullUrlLastCharPos: number = webFullUrl.length - 1;

    if (webFullUrl.length > 1 &&
      webFullUrl[webFullUrlLastCharPos] === '/') {
      webFullUrl = webFullUrl.substring(0, webFullUrlLastCharPos);
    }

    folderToEnsure = Utils.getWebRelativePath(webFullUrl, folderToEnsure);

    if (this.debug) {
      this.cmd.log(`folderToEnsure`);
      this.cmd.log(folderToEnsure);
      this.cmd.log('');
    }

    let nextFolder: string = '';
    let prevFolder: string = '';
    let folderIndex: number = 0;

    // build array of folders e.g. ["Shared%20Documents","22","54","55"]
    let folders: string[] = folderToEnsure.substring(1).split('/');

    if (this.debug) {
      this.cmd.log('folders to process');
      this.cmd.log(JSON.stringify(folders));
      this.cmd.log('');
    }

    // recursive function
    const checkOrAddFolder = (resolve: () => void, reject: (error: any) => void): void => {

      if (folderIndex === folders.length) {

        if (this.debug) {
          this.cmd.log(`All sub-folders exist`);
        }

        return resolve();
      }

      // append the next sub-folder to the folder path and check if it exists
      prevFolder = nextFolder;
      nextFolder += `/${folders[folderIndex]}`;
      const folderServerRelativeUrl = Utils.getServerRelativePath(webFullUrl, nextFolder);

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
        .catch(() => {
          const prevFolderServerRelativeUrl: string = Utils.getServerRelativePath(webFullUrl, prevFolder);
          const requestOptions: any = {
            url: `${webFullUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl=@a1)/AddSubFolderUsingPath(DecodedUrl=@a2)?@a1=%27${encodeURIComponent(prevFolderServerRelativeUrl)}%27&@a2=%27${encodeURIComponent(folders[folderIndex])}%27`,
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${siteAccessToken}`,
              'accept': 'application/json;odata=nometadata'
            }),
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