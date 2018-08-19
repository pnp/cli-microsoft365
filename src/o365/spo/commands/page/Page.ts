import * as request from 'request-promise-native';
import Utils from '../../../../Utils';
import { PageItem } from './PageItem';
import { ClientSidePage, CanvasSection, CanvasColumn, ClientSidePart } from './clientsidepages';

export class Page {
  public static getPage(name: string, webUrl: string, accessToken: string, cmd: CommandInstance, debug: boolean, verbose: boolean): Promise<ClientSidePage> {
    return new Promise((resolve: (page: ClientSidePage) => void, reject: (error: any) => void): void => {
      if (debug) {
        cmd.log(`Retrieved access token ${accessToken}`);
      }

      if (verbose) {
        cmd.log(`Retrieving information about the page...`);
      }

      let pageName: string = name;
      if (pageName.indexOf('.aspx') < 0) {
        pageName += '.aspx';
      }

      const requestOptions: any = {
        url: `${webUrl}/_api/web/getfilebyserverrelativeurl('${webUrl.substr(webUrl.indexOf('/', 8))}/SitePages/${encodeURIComponent(pageName)}')?$expand=ListItemAllFields/ClientSideApplicationId`,
        headers: Utils.getRequestHeaders({
          authorization: `Bearer ${accessToken}`,
          'content-type': 'application/json;charset=utf-8',
          accept: 'application/json;odata=nometadata'
        }),
        json: true
      };

      if (debug) {
        cmd.log('Executing web request...');
        cmd.log(requestOptions);
        cmd.log('');
      }

      request
        .get(requestOptions)
        .then((res: PageItem): void => {
          if (debug) {
            cmd.log('Response:');
            cmd.log(res);
            cmd.log('');
          }

          if (res.ListItemAllFields.ClientSideApplicationId !== 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec') {
            reject(`Page ${name} is not a modern page.`);
            return;
          }

          resolve(ClientSidePage.fromHtml(res.ListItemAllFields.CanvasContent1));
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
