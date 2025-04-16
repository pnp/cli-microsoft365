import request, { CliRequestOptions } from '../request.js';
import { formatting } from './formatting.js';

export interface SppModel {
  ConfidenceScore?: string;
  Explanations?: string;
  Schemas?: string;
  UniqueId: string;
  Publications?: any[];
  ModelSettings?: string;
}

export const spp = {
  /**
   * Asserts whether the specified site is a content center
   * @param siteUrl The URL of the site to check
   * @throws error when site is not a content center.
   */
  async assertSiteIsContentCenter(siteUrl: string): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${siteUrl}/_api/web?$select=WebTemplateConfiguration`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const response = await request.get<{ WebTemplateConfiguration: string }>(requestOptions);

    if (response.WebTemplateConfiguration !== 'CONTENTCTR#0') {
      throw Error(`${siteUrl} is not a content site.`);
    }
  },
  /**
   * Gets a SharePoint Premium model by title or unique id
   * @param contentCenterUrl a content center site URL
   * @param title model title
   * @param id model unique id
   * @returns SharePoint Premium model
   */

  async getModel(contentCenterUrl: string, title?: string, id?: string): Promise<SppModel> {
    let requestUrl = `${contentCenterUrl}/_api/machinelearning/models/`;

    if (title) {
      let requestTitle = title.toLowerCase();

      if (!requestTitle.endsWith('.classifier')) {
        requestTitle += '.classifier';
      }

      requestUrl += `getbytitle('${formatting.encodeQueryParameter(requestTitle)}')`;
    }
    else {
      requestUrl += `getbyuniqueid('${id}')`;
    }

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const result = await request.get<SppModel>(requestOptions);

    if ((result as any)['odata.null'] === true) {
      throw Error("Model not found.");
    }

    return result;
  }
};