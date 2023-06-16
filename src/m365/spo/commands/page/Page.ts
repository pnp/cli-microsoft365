import { Logger } from '../../../../cli/Logger';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { urlUtil } from '../../../../utils/urlUtil';
import { ClientSidePageProperties } from './ClientSidePageProperties';
import { CanvasColumn, CanvasSection, ClientSidePage, ClientSidePart } from './clientsidepages';
import { PageItem } from './PageItem';
import { getControlTypeDisplayName } from './pageMethods';

export const supportedPageLayouts = ['Article', 'Home', 'SingleWebPartAppPage', 'RepostPage', 'HeaderlessSearchResults', 'Spaces', 'Topic'];
export const supportedPromoteAs = ['HomePage', 'NewsPage', 'Template'];

export class Page {
  public static async getPage(name: string, webUrl: string, logger: Logger, debug: boolean, verbose: boolean): Promise<ClientSidePage> {
    if (verbose) {
      logger.logToStderr(`Retrieving information about the page...`);
    }

    const pageName: string = this.getPageNameWithExtension(name);

    const requestOptions: any = {
      url: `${webUrl}/_api/web/getfilebyserverrelativeurl('${urlUtil.getServerRelativeSiteUrl(webUrl)}/SitePages/${formatting.encodeQueryParameter(pageName)}')?$expand=ListItemAllFields/ClientSideApplicationId`,
      headers: {
        'content-type': 'application/json;charset=utf-8',
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const res: PageItem = await request.get<PageItem>(requestOptions);
    if (res.ListItemAllFields.ClientSideApplicationId !== 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec') {
      throw `Page ${name} is not a modern page.`;
    }

    return ClientSidePage.fromHtml(res.ListItemAllFields.CanvasContent1);
  }

  public static async checkout(name: string, webUrl: string, logger: Logger, debug: boolean, verbose: boolean): Promise<ClientSidePageProperties> {
    if (verbose) {
      logger.log(`Checking out ${name} page...`);
    }

    const pageName: string = this.getPageNameWithExtension(name);
    const requestOptions: any = {
      url: `${webUrl}/_api/sitepages/pages/GetByUrl('sitepages/${formatting.encodeQueryParameter(pageName)}')/checkoutpage`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const pageData: ClientSidePageProperties = await request.post<ClientSidePageProperties>(requestOptions);
    if (!pageData) {
      throw `Page ${name} information not retrieved with the checkout`;
    }

    if (verbose) {
      logger.log(`Page ${name} is now checked out`);
    }

    return pageData;
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

  public static getColumnsInformation(column: CanvasColumn, isJSONOutput: boolean): any {
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
    };
  }

  private static getPageNameWithExtension(name: string): string {
    let pageName: string = name;
    if (pageName.indexOf('.aspx') < 0) {
      pageName += '.aspx';
    }

    return pageName;
  }
}
