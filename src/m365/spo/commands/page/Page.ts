import request from '../../../../request';
import { PageItem } from './PageItem';
import { ClientSidePage, CanvasSection, CanvasColumn, ClientSidePart } from './clientsidepages';
import Utils from '../../../../Utils';
import { CommandInstance } from '../../../../cli';

export class Page {
  public static getPage(name: string, webUrl: string, cmd: CommandInstance, debug: boolean, verbose: boolean): Promise<ClientSidePage> {
    return new Promise((resolve: (page: ClientSidePage) => void, reject: (error: any) => void): void => {
      if (verbose) {
        cmd.log(`Retrieving information about the page...`);
      }

      let pageName: string = name;
      if (pageName.indexOf('.aspx') < 0) {
        pageName += '.aspx';
      }

      const requestOptions: any = {
        url: `${webUrl}/_api/web/getfilebyserverrelativeurl('${Utils.getServerRelativeSiteUrl(webUrl)}/SitePages/${encodeURIComponent(pageName)}')?$expand=ListItemAllFields/ClientSideApplicationId`,
        headers: {
          'content-type': 'application/json;charset=utf-8',
          accept: 'application/json;odata=nometadata'
        },
        json: true
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

  public static getControlTypeDisplayName(controlType: number): string {
    switch (controlType) {
      case 0:
        return 'Empty column';
      case 3:
        return 'Client-side web part';
      case 4:
        return 'Client-side text';
      default:
        return '' + controlType;
    }
  }

  public static getControlsInformation(control: ClientSidePart, isJSONOutput: boolean): ClientSidePart {
    // remove the column property to be able to serialize the object to JSON
    delete control.column;

    if (!isJSONOutput) {
      (control as any).controlType = this.getControlTypeDisplayName((control as any).controlType);
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
}
