import AzmgmtCommand from "./AzmgmtCommand";
import request from '../../request';
import { AzmgmtResponse } from './AzmgmtResponse';

export abstract class AzmgmtItemsListCommand<T> extends AzmgmtCommand {
  protected items: T[];

  /* istanbul ignore next */
  constructor() {
    super();
    this.items = [];
  }

  protected getAllItems(url: string, cmd: CommandInstance, firstRun: boolean): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      const requestOptions: any = {
        url: url,
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
            this
              .getAllItems(res.nextLink as string, cmd, false)
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