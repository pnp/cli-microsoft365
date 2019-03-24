import auth from '../GraphAuth';
import GraphCommand from "../GraphCommand";
import request from '../../../request';
import { GraphResponse } from '../GraphResponse';

export abstract class GraphItemsListCommand<T> extends GraphCommand {
  protected items: T[];

  /* istanbul ignore next */
  constructor() {
    super();
    this.items = [];
  }

  protected getAllItems(url: string, cmd: CommandInstance, firstRun: boolean): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      auth
        .ensureAccessToken(auth.service.resource, cmd, this.debug)
        .then((): Promise<GraphResponse<T>> => {
          const requestOptions: any = {
            url: url,
            headers: {
              authorization: `Bearer ${auth.service.accessToken}`,
              accept: 'application/json;odata.metadata=none'
            },
            json: true
          };

          return request.get(requestOptions);
        })
        .then((res: GraphResponse<T>): void => {
          if (firstRun) {
            this.items = [];
          }

          this.items = this.items.concat(res.value);

          if (res['@odata.nextLink']) {
            this
              .getAllItems(res['@odata.nextLink'] as string, cmd, false)
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