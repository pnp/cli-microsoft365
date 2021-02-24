import { Logger } from '../../../../cli';
import request from '../../../../request';
import Utils from '../../../../Utils';
import { ClientSidePageProperties } from './ClientSidePageProperties';
import { CanvasColumn, CanvasSection, ClientSidePage, ClientSidePart } from './clientsidepages';
import { PageItem } from './PageItem';
import { getControlTypeDisplayName } from './pageMethods';

export class Page {
  public static getPage(name: string, webUrl: string, logger: Logger, debug: boolean, verbose: boolean): Promise<ClientSidePage> {
    return new Promise((resolve: (page: ClientSidePage) => void, reject: (error: any) => void): void => {
      if (verbose) {
        logger.logToStderr(`Retrieving information about the page...`);
      }

      let pageName: string = this.getPageNameWithExtension(name);

      const requestOptions: any = {
        url: `${webUrl}/_api/web/getfilebyserverrelativeurl('${Utils.getServerRelativeSiteUrl(webUrl)}/SitePages/${encodeURIComponent(pageName)}')?$expand=ListItemAllFields/ClientSideApplicationId`,
        headers: {
          'content-type': 'application/json;charset=utf-8',
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      request
        .get<PageItem>(requestOptions)
        .then((res: PageItem): void => {
          if (res.ListItemAllFields.ClientSideApplicationId !== 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec') {
            reject(`Page ${name} is not a modern page.`);
            return;
          }

          try {
            resolve(ClientSidePage.fromHtml(res.ListItemAllFields.CanvasContent1));
          }
          catch (e) {
            reject(e);
          }
        }, (error: any): void => {
          reject(error);
        });
    });
  }

  public static checkout(name: string, webUrl: string, logger: Logger, debug: boolean, verbose: boolean): Promise<ClientSidePageProperties> {
    return new Promise<ClientSidePageProperties>((resolve: (page: ClientSidePageProperties) => void, reject: (error: any) => void): void => {
      if (verbose) {
        logger.log(`Checking out ${name} page...`);
      }

      const pageName: string = this.getPageNameWithExtension(name);
      const requestOptions: any = {
        url: `${webUrl}/_api/sitepages/pages/GetByUrl('sitepages/${encodeURIComponent(pageName)}')/checkoutpage`,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      request
        .post<ClientSidePageProperties>(requestOptions)
        .then((pageData: ClientSidePageProperties) => {
          if (!pageData) {
            reject(`Page ${name} information not retrieved with the checkout`);
            return;
          }

          if (verbose) {
            logger.log(`Page ${name} is now checked out`);
          }

          resolve(pageData);
        }, (error: any): void => {
          reject(error);
        });
    });
  }

  public static save(name: string, webUrl: string, canvasContent: any, logger: Logger, debug: boolean, verbose: boolean): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      if (verbose) {
        logger.log(`Saving ${name} page...`);
      }

      if (!canvasContent) {
        reject('No canvas content was provided');
        return;
      }

      const pageName: string = this.getPageNameWithExtension(name);
      const requestOptions: any = {
        url: `${webUrl}/_api/sitepages/pages/GetByUrl('sitepages/${encodeURIComponent(pageName)}')/savepage`,
        headers: {
          'accept': 'application/json;odata=nometadata',
          'content-type': 'application/json;odata=nometadata'
        },
        data: {
          CanvasContent1: JSON.stringify(canvasContent)
        },
        responseType: 'json'
      };

      request
        .post(requestOptions)
        .then((res: any) => {
          if (verbose) {
            logger.log(res);
          }

          resolve();
        }, (error: any): void => {
          reject(error);
        });
    });
  }

  public static getControlsInformation(control: ClientSidePart, isJSONOutput: boolean): ClientSidePart {
    // remove the column property to be able to serialize the object to JSON
    delete control.column;

    if (!isJSONOutput) {
      (control as any).controlType = getControlTypeDisplayName((control as any).controlType);
    }

    if (!control.dynamicDataPaths) {
      delete control.dynamicDataPaths;
    }

    if (!control.dynamicDataValues) {
      delete control.dynamicDataValues;
    }

    return control;
  }

  public static getColumnsInformation(column: CanvasColumn, isJSONOutput: boolean) {
    const output: any = {
      factor: column.factor,
      order: column.order
    };

    if (isJSONOutput) {
      output.dataVersion = column.dataVersion;
      output.jsonData = column.jsonData;
    }

    return output;
  }

  public static getSectionInformation(section: CanvasSection, isJSONOutput: boolean): any {
    return {
      order: section.order,
      columns: section.columns.map(column => this.getColumnsInformation(column, isJSONOutput))
    }
  }

  private static getPageNameWithExtension(name: string): string {
    let pageName: string = name;
    if (pageName.indexOf('.aspx') < 0) {
      pageName += '.aspx';
    }

    return pageName;
  }
}
