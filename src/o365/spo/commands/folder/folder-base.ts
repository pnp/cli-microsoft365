import SpoCommand from "../../SpoCommand";
import * as url from 'url';

export abstract class FolderBaseCommand extends SpoCommand {

  public formatRelativeUrl(relativeUrl: string): string {

    // add '/' at 0
    if (relativeUrl.charAt(0) !== '/') {
      relativeUrl = `/${relativeUrl}`;
    }

    // remove last '/' of url
    if (relativeUrl.lastIndexOf('/') === relativeUrl.length - 1) {
      relativeUrl = relativeUrl.substring(0, relativeUrl.length - 1);
    }

    return relativeUrl;
  }

  public getWebRelativeUrlFromWebUrl(webUrl: string): string {

    const tenantUrl = `${url.parse(webUrl).protocol}//${url.parse(webUrl).hostname}`;
    let webRelativeUrl = webUrl.replace(tenantUrl, '');

    return this.formatRelativeUrl(webRelativeUrl);
  }
}