import request, { CliRequestOptions } from '../request.js';

export interface ContentCenterSite {
  WebTemplateConfiguration: string;
}

export interface SppModel {
  AIBuilderHybridModelType: string;
  AzureCognitivePrebuiltModelName: string;
  BaseContentTypeName: string;
  ConfidenceScore: string;
  ContentTypeGroup: string;
  ContentTypeId: string;
  ContentTypeName: string;
  Created: string;
  CreatedBy: string;
  DriveId: string;
  Explanations: string;
  ID: number,
  LastTrained: string;
  ListID: string;
  ModelSettings: string;
  ModelType: number,
  Modified: string;
  ModifiedBy: string;
  ObjectId: string;
  PublicationType: number;
  Schemas: string;
  SourceSiteUrl: string;
  SourceUrl: string;
  SourceWebServerRelativeUrl: string;
  UniqueId: string;
}

export const spp = {
  async isContentCenter(siteUrl: string): Promise<boolean> {
    const requestOptions: CliRequestOptions = {
      url: `${siteUrl}/_api/web?$select=WebTemplateConfiguration`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const response = await request.get<ContentCenterSite>(requestOptions);
    return response.WebTemplateConfiguration === 'CONTENTCTR#0';
  }
};