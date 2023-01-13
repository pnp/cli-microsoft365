import request, { CliRequestOptions } from "../../../../request";
import { formatting } from "../../../../utils/formatting";
import { urlUtil } from "../../../../utils/urlUtil";
import { GraphFileDetails } from "./GraphFileDetails";

export class FileSharingLinkUtil {
  public static readonly allowedScopes: string[] = ['anonymous', 'users', 'organization'];

  public static async getFileDetails(webUrl: string, fileId?: string, fileUrl?: string): Promise<GraphFileDetails> {
    let requestUrl: string = `${webUrl}/_api/web/`;

    if (fileUrl) {
      const fileServerRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, fileUrl);
      requestUrl += `GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(fileServerRelativeUrl)}')`;
    }
    else {
      requestUrl += `GetFileById('${fileId}')`;
    }

    requestUrl += '?$select=SiteId,VroomItemId,VroomDriveId';

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const res = await request.get<GraphFileDetails>(requestOptions);
    return res;
  }

}