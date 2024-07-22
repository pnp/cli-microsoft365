import request, { CliRequestOptions } from '../request.js';

export const spp = {
  /**
   * Asserts whether the specified site is a content center
   * @param siteUrl The URL of the site to check
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
  }
};