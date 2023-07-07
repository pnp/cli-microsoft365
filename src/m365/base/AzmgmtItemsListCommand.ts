import * as url from 'url';
import { Logger } from '../../cli/Logger';
import request, { CliRequestOptions } from '../../request';
import { ODataResponse } from '../../utils/odata';
import AzmgmtCommand from "./AzmgmtCommand";

export abstract class AzmgmtItemsListCommand<T> extends AzmgmtCommand {
  protected items: T[];

  constructor() {
    super();
    this.items = [];
  }

  protected async getAllItems(_url: string, logger: Logger, firstRun: boolean): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: _url,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const res = await request.get<ODataResponse<T>>(requestOptions);

    if (firstRun) {
      this.items = [];
    }

    this.items = this.items.concat(res.value);

    if (res.nextLink) {
      // when retrieving Flows as admin, the API returns nextLink
      // pointing to https://emea.api.flow.microsoft.com:11777
      // which leads to authentication exceptions because it's not an AAD
      // resource for which we can retrieve an access token, so we need to
      // rewrite it back to the API management URL
      const nextLinkUrl: url.URL = new url.URL(res.nextLink);
      if (nextLinkUrl.host !== 'management.azure.com') {
        nextLinkUrl.host = 'management.azure.com';
        nextLinkUrl.port = '';
      }

      const nextLink: string = nextLinkUrl.href;
      await this.getAllItems(nextLink, logger, false);
    }
    else {
      return;
    }
  }
}