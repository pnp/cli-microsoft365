import { Logger } from "../../cli";
import request from '../../request';
import GraphCommand from "./GraphCommand";
import { GraphResponse } from './GraphResponse';

export abstract class GraphItemsListCommand<T> extends GraphCommand {
  protected items: T[];

  constructor() {
    super();
    this.items = [];
  }

  protected getAllItems(url: string, logger: Logger, firstRun: boolean): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      const requestOptions: any = {
        url: url,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      request
        .get<GraphResponse<T>>(requestOptions)
        .then((res: GraphResponse<T>): void => {
          if (firstRun) {
            this.items = [];
          }

          this.items = this.items.concat(res.value);

          if (res['@odata.nextLink']) {
            this
              .getAllItems(res['@odata.nextLink'] as string, logger, false)
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