export interface PageHeader {
  description: string;
  id: string;
  instanceId: string;
  title: string;
  serverProcessedContent: PageHeaderServerProcessedContent,
  dataVersion: string;
  properties: PageHeaderProperties
}

export interface CustomPageHeader extends PageHeader {
  properties: CustomPageHeaderProperties;
  serverProcessedContent: CustomPageHeaderServerProcessedContent,
}

interface PageHeaderProperties {
  imageSourceType: number;
  topicHeader: string;
  layoutType: "NoImage" | "FullWidthImage" | "ColorBlock" | "CutInShape";
  showTopicHeader: boolean;
  showPublishDate: boolean;
  textAlignment: "Left" | "Center";
  title: string;
}

export interface CustomPageHeaderProperties extends PageHeaderProperties {
  altText: string;
  authors: string[];
  listId: string;
  siteId: string;
  translateX: number;
  translateY: number;
  uniqueId: string;
  webId: string;
}

interface PageHeaderServerProcessedContent {
  htmlStrings: any,
  imageSources: {
    imageSource?: string;
  };
  links: any;
  searchablePlainTexts: any;
}

export interface CustomPageHeaderServerProcessedContent extends PageHeaderServerProcessedContent {
  customMetadata: {
    imageSource: {
      listId: string;
      siteId: string;
      uniqueId: string;
      webId: string;
    }
  }
}