import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ClientSidePageProperties } from './ClientSidePageProperties';
import { CustomPageHeader, CustomPageHeaderProperties, CustomPageHeaderServerProcessedContent, PageHeader } from './PageHeader';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  altText?: string;
  authors?: string;
  imageUrl?: string;
  topicHeader?: string;
  layout?: string;
  pageName: string;
  showTopicHeader?: boolean;
  showPublishDate?: boolean;
  textAlignment?: string;
  translateX?: number;
  translateY?: number;
  type?: string;
  webUrl: string;
}

class SpoPageHeaderSetCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_HEADER_SET;
  }

  public get description(): string {
    return 'Sets modern page header';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        altText: typeof args.options.altText !== 'undefined',
        authors: typeof args.options.authors !== 'undefined',
        imageUrl: typeof args.options.imageUrl !== 'undefined',
        topicHeader: typeof args.options.topicHeader !== 'undefined',
        layout: args.options.layout,
        showTopicHeader: args.options.showTopicHeader,
        showPublishDate: args.options.showPublishDate,
        textAlignment: args.options.textAlignment,
        translateX: typeof args.options.translateX !== 'undefined',
        translateY: typeof args.options.translateY !== 'undefined',
        type: args.options.type
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --pageName <pageName>'
      },
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-t, --type [type]',
        autocomplete: ['None', 'Default', 'Custom']
      },
      {
        option: '--imageUrl [imageUrl]'
      },
      {
        option: '--altText [altText]'
      },
      {
        option: '-x, --translateX [translateX]'
      },
      {
        option: '-y, --translateY [translateY]'
      },
      {
        option: '--layout [layout]',
        autocomplete: ['FullWidthImage', 'NoImage', 'ColorBlock', 'CutInShape']
      },
      {
        option: '--textAlignment [textAlignment]',
        autocomplete: ['Left', 'Center']
      },
      {
        option: '--showTopicHeader'
      },
      {
        option: '--showPublishDate'
      },
      {
        option: '--topicHeader [topicHeader]'
      },
      {
        option: '--authors [authors]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.type &&
          args.options.type !== 'None' &&
          args.options.type !== 'Default' &&
          args.options.type !== 'Custom') {
          return `${args.options.type} is not a valid type value. Allowed values None|Default|Custom`;
        }

        if (args.options.translateX && isNaN(args.options.translateX)) {
          return `${args.options.translateX} is not a valid number`;
        }

        if (args.options.translateY && isNaN(args.options.translateY)) {
          return `${args.options.translateY} is not a valid number`;
        }

        if (args.options.layout &&
          args.options.layout !== 'FullWidthImage' &&
          args.options.layout !== 'NoImage' &&
          args.options.layout !== 'ColorBlock' &&
          args.options.layout !== 'CutInShape') {
          return `${args.options.layout} is not a valid layout value. Allowed values FullWidthImage|NoImage|ColorBlock|CutInShape`;
        }

        if (args.options.textAlignment &&
          args.options.textAlignment !== 'Left' &&
          args.options.textAlignment !== 'Center') {
          return `${args.options.textAlignment} is not a valid textAlignment value. Allowed values Left|Center`;
        }

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  protected getExcludedOptionsWithUrls(): string[] | undefined {
    return ['imageUrl'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const noPageHeader: PageHeader = {
      "id": "cbe7b0a9-3504-44dd-a3a3-0e5cacd07788",
      "instanceId": "cbe7b0a9-3504-44dd-a3a3-0e5cacd07788",
      "title": "Title Region",
      "description": "Title Region Description",
      "serverProcessedContent": {
        "htmlStrings": {},
        "searchablePlainTexts": {},
        "imageSources": {},
        "links": {}
      },
      "dataVersion": "1.4",
      "properties": {
        "title": "",
        "imageSourceType": 4,
        "layoutType": "NoImage",
        "textAlignment": "Left",
        "showTopicHeader": false,
        "showPublishDate": false,
        "topicHeader": ""
      }
    };
    const defaultPageHeader: PageHeader = {
      "id": "cbe7b0a9-3504-44dd-a3a3-0e5cacd07788",
      "instanceId": "cbe7b0a9-3504-44dd-a3a3-0e5cacd07788",
      "title": "Title Region",
      "description": "Title Region Description",
      "serverProcessedContent": {
        "htmlStrings": {},
        "searchablePlainTexts": {},
        "imageSources": {},
        "links": {}
      },
      "dataVersion": "1.4",
      "properties": {
        "title": "",
        "imageSourceType": 4,
        "layoutType": "FullWidthImage",
        "textAlignment": "Left",
        "showTopicHeader": false,
        "showPublishDate": false,
        "topicHeader": ""
      }
    };
    const customPageHeader: CustomPageHeader = {
      "id": "cbe7b0a9-3504-44dd-a3a3-0e5cacd07788",
      "instanceId": "cbe7b0a9-3504-44dd-a3a3-0e5cacd07788",
      "title": "Title Region",
      "description": "Title Region Description",
      "serverProcessedContent": {
        "htmlStrings": {},
        "searchablePlainTexts": {},
        "imageSources": {
          "imageSource": ""
        },
        "links": {},
        "customMetadata": {
          "imageSource": {
            "siteId": "",
            "webId": "",
            "listId": "",
            "uniqueId": ""
          }
        }
      },
      "dataVersion": "1.4",
      "properties": {
        "title": "",
        "imageSourceType": 2,
        "layoutType": "FullWidthImage",
        "textAlignment": "Left",
        "showTopicHeader": false,
        "showPublishDate": false,
        "topicHeader": "",
        "authors": [],
        "altText": "",
        "webId": "",
        "siteId": "",
        "listId": "",
        "uniqueId": "",
        "translateX": 0,
        "translateY": 0
      }
    };
    let header: PageHeader | CustomPageHeader = defaultPageHeader;
    let pageFullName: string = args.options.pageName.toLowerCase();
    if (pageFullName.indexOf('.aspx') < 0) {
      pageFullName += '.aspx';
    }

    let canvasContent: string = "";
    let bannerImageUrl: string = "";
    let description: string = "";
    let title: string = "";
    let authorByline: string[] = args.options.authors ? args.options.authors.split(',').map(a => a.trim()) : [];
    let topicHeader: string = args.options.topicHeader || "";

    if (this.verbose) {
      logger.logToStderr(`Retrieving information about the page...`);
    }

    try {
      let requestOptions: any = {
        url: `${args.options.webUrl}/_api/sitepages/pages/GetByUrl('sitepages/${formatting.encodeQueryParameter(pageFullName)}')?$select=IsPageCheckedOutToCurrentUser,Title`,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const page = await request.get<{ IsPageCheckedOutToCurrentUser: boolean, Title: string; }>(requestOptions);
      title = page.Title;

      let pageData: ClientSidePageProperties;
      if (page.IsPageCheckedOutToCurrentUser) {
        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/sitepages/pages/GetByUrl('sitepages/${formatting.encodeQueryParameter(pageFullName)}')?$expand=ListItemAllFields`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        pageData = await request.get<ClientSidePageProperties>(requestOptions);
      }
      else {
        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/sitepages/pages/GetByUrl('sitepages/${formatting.encodeQueryParameter(pageFullName)}')/checkoutpage`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        pageData = await request.post<ClientSidePageProperties>(requestOptions);
      }

      switch (args.options.type) {
        case 'None':
          header = noPageHeader;
          break;
        case 'Default':
          header = defaultPageHeader;
          break;
        case 'Custom':
          header = customPageHeader;
          break;
        default:
          header = defaultPageHeader;
      }

      if (pageData) {
        canvasContent = pageData.CanvasContent1;
        authorByline = authorByline.length > 0 ? authorByline : pageData.AuthorByline;
        bannerImageUrl = pageData.BannerImageUrl;
        description = pageData.Description;
        title = pageData.Title;
        topicHeader = topicHeader || pageData.TopicHeader || "";
      }

      header.properties.title = title;
      header.properties.textAlignment = args.options.textAlignment as any || 'Left';
      header.properties.showTopicHeader = args.options.showTopicHeader || false;
      header.properties.topicHeader = args.options.topicHeader || '';
      header.properties.showPublishDate = args.options.showPublishDate || false;

      if (args.options.type !== 'None') {
        header.properties.layoutType = args.options.layout as any || 'FullWidthImage';
      }

      if (args.options.type === 'Custom') {
        header.serverProcessedContent.imageSources = {
          imageSource: args.options.imageUrl || ''
        };
        const properties: CustomPageHeaderProperties = header.properties as CustomPageHeaderProperties;
        properties.altText = args.options.altText || '';
        properties.translateX = args.options.translateX || 0;
        properties.translateY = args.options.translateY || 0;
        header.properties = properties;

        if (!args.options.imageUrl) {
          (header.serverProcessedContent as CustomPageHeaderServerProcessedContent).customMetadata = {
            imageSource: {
              siteId: '',
              webId: '',
              listId: '',
              uniqueId: ''
            }
          };
          properties.listId = '';
          properties.siteId = '';
          properties.uniqueId = '';
          properties.webId = '';
          header.properties = properties;
        }
        else {
          const res = await Promise.all([
            this.getSiteId(args.options.webUrl, this.verbose, logger),
            this.getWebId(args.options.webUrl, this.verbose, logger),
            this.getImageInfo(args.options.webUrl, args.options.imageUrl as string, this.verbose, logger)
          ]);

          (header.serverProcessedContent as CustomPageHeaderServerProcessedContent).customMetadata = {
            imageSource: {
              siteId: res[0].Id,
              webId: res[1].Id,
              listId: res[2].ListId,
              uniqueId: res[2].UniqueId
            }
          };
          const properties: CustomPageHeaderProperties = header.properties as CustomPageHeaderProperties;
          properties.listId = res[2].ListId;
          properties.siteId = res[0].Id;
          properties.uniqueId = res[2].UniqueId;
          properties.webId = res[1].Id;
          header.properties = properties;
        }
      }

      const requestBody: any = {
        LayoutWebpartsContent: JSON.stringify([header])
      };

      if (title) {
        requestBody.Title = title;
      }
      if (topicHeader) {
        requestBody.TopicHeader = topicHeader;
      }
      if (description) {
        requestBody.Description = description;
      }
      if (authorByline) {
        requestBody.AuthorByline = authorByline;
      }
      if (bannerImageUrl) {
        requestBody.BannerImageUrl = bannerImageUrl;
      }
      if (canvasContent) {
        requestBody.CanvasContent1 = canvasContent;
      }

      requestOptions = {
        url: `${args.options.webUrl}/_api/sitepages/pages/GetByUrl('sitepages/${formatting.encodeQueryParameter(pageFullName)}')/SavePageAsDraft`,
        headers: {
          'X-HTTP-Method': 'MERGE',
          'IF-MATCH': '*',
          'content-type': 'application/json;odata=nometadata',
          accept: 'application/json;odata=nometadata'
        },
        data: requestBody,
        responseType: 'json'
      };

      return request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getSiteId(siteUrl: string, verbose: boolean, logger: Logger): Promise<any> {
    if (verbose) {
      logger.logToStderr(`Retrieving information about the site collection...`);
    }

    const requestOptions: any = {
      url: `${siteUrl}/_api/site?$select=Id`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request.get(requestOptions);
  }

  private getWebId(siteUrl: string, verbose: boolean, logger: Logger): Promise<any> {
    if (verbose) {
      logger.logToStderr(`Retrieving information about the site...`);
    }

    const requestOptions: any = {
      url: `${siteUrl}/_api/web?$select=Id`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request.get(requestOptions);
  }

  private getImageInfo(siteUrl: string, imageUrl: string, verbose: boolean, logger: Logger): Promise<any> {
    if (verbose) {
      logger.logToStderr(`Retrieving information about the header image...`);
    }

    const requestOptions: any = {
      url: `${siteUrl}/_api/web/getfilebyserverrelativeurl('${formatting.encodeQueryParameter(imageUrl)}')?$select=ListId,UniqueId`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request.get(requestOptions);
  }
}

module.exports = new SpoPageHeaderSetCommand();
