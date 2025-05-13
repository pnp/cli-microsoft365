import { Logger } from '../cli/Logger.js';
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
   * @param logger Logger instance
   * @param verbose Whether to log verbose messages
   * @throws error when site is not a content center.
   */
  async assertSiteIsContentCenter(siteUrl: string, logger: Logger, verbose: boolean): Promise<void> {
    if (verbose) {
      await logger.log(`Asserting if ${siteUrl} is a valid content center...`);
    }

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
   * Gets a SharePoint Premium model by title
   * @param contentCenterUrl a content center site URL
   * @param title model title
   * @param logger Logger instance
   * @param verbose Whether to log verbose messages
   * @returns SharePoint Premium model
   */
  async getModelByTitle(contentCenterUrl: string, title: string, logger: Logger, verbose: boolean): Promise<SppModel> {
    if (verbose) {
      await logger.log(`Retrieving model information...`);
    }

    let requestTitle = title.toLowerCase();

    if (!requestTitle.endsWith('.classifier')) {
      requestTitle += '.classifier';
    }

    const requestUrl = `${contentCenterUrl}/_api/machinelearning/models/getbytitle('${formatting.encodeQueryParameter(requestTitle)}')`;
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
  },
  /**
   * Gets a SharePoint Premium model by unique id
   * @param contentCenterUrl a content center site URL
   * @param id model unique id
   * @param logger Logger instance
   * @param verbose Whether to log verbose messages
   * @returns SharePoint Premium model
   */
  async getModelById(contentCenterUrl: string, id: string, logger: Logger, verbose: boolean): Promise<SppModel> {
    if (verbose) {
      await logger.log(`Retrieving model information from...`);
    }

    const requestUrl = `${contentCenterUrl}/_api/machinelearning/models/getbyuniqueid('${id}')`;

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return await request.get<SppModel>(requestOptions);
  }
};