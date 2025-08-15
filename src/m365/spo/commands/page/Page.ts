import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { ClientSidePageProperties } from './ClientSidePageProperties.js';
import { CanvasColumn, CanvasSection, ClientSidePage, ClientSidePart } from './clientsidepages.js';
import { PageItem } from './PageItem.js';
import { getControlTypeDisplayName } from './pageMethods.js';

export const supportedPageLayouts = ['Article', 'Home', 'SingleWebPartAppPage', 'RepostPage', 'HeaderlessSearchResults', 'Spaces', 'Topic'];
export const supportedPromoteAs = ['HomePage', 'NewsPage', 'Template'];

export class Page {
  public static async getPage(name: string, webUrl: string, logger: Logger, debug: boolean, verbose: boolean): Promise<ClientSidePage> {
    if (verbose) {
      await logger.logToStderr(`Retrieving information about the page...`);
    }

    const pageName: string = this.getPageNameWithExtension(name);

    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${urlUtil.getServerRelativeSiteUrl(webUrl)}/SitePages/${formatting.encodeQueryParameter(pageName)}')?$expand=ListItemAllFields/ClientSideApplicationId`,
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

  public static async checkout(name: string, webUrl: string, logger: Logger, verbose: boolean): Promise<ClientSidePageProperties> {
    if (verbose) {
      await logger.log(`Checking out ${name} page...`);
    }

    const pageName: string = this.getPageNameWithExtension(name);
    const requestOptions: CliRequestOptions = {
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
      await logger.log(`Page ${name} is now checked out`);
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
    const sectionOutput: any = {
      order: section.order
    };

    if (this.isVerticalSection(section)) {
      sectionOutput.isVertical = true;
    }

    sectionOutput.columns = section.columns.map(column => this.getColumnsInformation(column, isJSONOutput));

    return sectionOutput;
  }

  /**
   * Publish a modern page in SharePoint Online
   * @param webUrl Absolute URL of the SharePoint site where the page is located
   * @param pageName List relative url of the page to publish
   */
  public static async publishPage(webUrl: string, pageName: string): Promise<void> {
    const filePath = `${urlUtil.getServerRelativeSiteUrl(webUrl)}/SitePages/${pageName}`;
    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(filePath)}')/Publish()`,
      headers: {
        accept: 'application/json;odata=nometadata'
      }
    };

    await request.post(requestOptions);
  }

  private static getPageNameWithExtension(name: string): string {
    let pageName: string = name;
    if (pageName.indexOf('.aspx') < 0) {
      pageName += '.aspx';
    }

    return pageName;
  }

  private static isVerticalSection(section: CanvasSection): boolean {
    return section.layoutIndex === 2 && section?.controlData?.position?.sectionFactor === 12;
  }
}
