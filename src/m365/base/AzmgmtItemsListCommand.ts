import * as url from 'url';
import { Logger } from "../../cli";
import request from '../../request';
import AzmgmtCommand from "./AzmgmtCommand";
import { AzmgmtResponse } from './AzmgmtResponse';

export abstract class AzmgmtItemsListCommand<T> extends AzmgmtCommand {
  protected items: T[];

  constructor() {
    super();
    this.items = [];
  }

  protected getAllItems(_url: string, logger: Logger, firstRun: boolean): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      const requestOptions: any = {
        url: _url,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        json: true
      };

      request
        .get<AzmgmtResponse<T>>(requestOptions)
        .then((res: AzmgmtResponse<T>): void => {
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

            this
              .getAllItems(nextLink, logger, false)
              .then((): void => {
                resolve();
              }, (err: any): void => {
                reject(err);
              });
          }
          else {
            resolve();
          }
        }, (err: any): void => {
          reject(err);
        });
    });
  }
}