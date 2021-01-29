import { isNumber } from 'util';
import { v4 } from 'uuid';
import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { StandardWebPart, StandardWebPartUtils } from '../../StandardWebPartTypes';
import { Control } from './canvasContent';
import { ClientSidePageProperties } from './ClientSidePageProperties';
import {
  ClientSidePageComponent, ClientSideWebpart
} from './clientsidepages';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  pageName: string;
  webUrl: string;
  standardWebPart?: StandardWebPart;
  webPartData?: string;
  webPartId?: string;
  webPartProperties?: string;
  section?: number;
  column?: number;
  order?: number;
}

class SpoPageClientSideWebPartAddCommand extends SpoCommand {
  public get name(): string {
    return `${commands.PAGE_CLIENTSIDEWEBPART_ADD}`;
  }

  public get description(): string {
    return 'Adds a client-side web part to a modern page';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.standardWebPart = args.options.standardWebPart;
    telemetryProps.webPartData = typeof args.options.webPartData !== 'undefined';
    telemetryProps.webPartId = typeof args.options.webPartId !== 'undefined';
    telemetryProps.webPartProperties = typeof args.options.webPartProperties !== 'undefined';
    telemetryProps.section = typeof args.options.section !== 'undefined';
    telemetryProps.column = typeof args.options.column !== 'undefined';
    telemetryProps.order = typeof args.options.order !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    let canvasContent: Control[];

    let layoutWebpartsContent: string = "";
    let authorByline: string[] = [""];
    let bannerImageUrl: string = "";
    let description: string = "";
    let title: string = "";
    let topicHeader: string = "";
    let pageData: any = {};

    let pageFullName: string = args.options.pageName;
    if (args.options.pageName.indexOf('.aspx') < 0) {
      pageFullName += '.aspx';
    }

    if (this.verbose) {
      logger.logToStderr(`Retrieving page information...`);
    }

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/sitepages/pages/GetByUrl('sitepages/${encodeURIComponent(pageFullName)}')`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get<ClientSidePageProperties>(requestOptions)
      .then((res: ClientSidePageProperties): Promise<ClientSidePageProperties> => {
        if (res.IsPageCheckedOutToCurrentUser) {
          return Promise.resolve(res);
        }

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/sitepages/pages/GetByUrl('sitepages/${encodeURIComponent(pageFullName)}')/checkoutpage`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        return request.post<ClientSidePageProperties>(requestOptions);
      })
      .then((res: ClientSidePageProperties): Promise<ClientSideWebpart> => {
        if (res) {
          layoutWebpartsContent = res.LayoutWebpartsContent;
          authorByline = res.AuthorByline;
          bannerImageUrl = res.BannerImageUrl;
          description = res.Description;
          title = res.Title;
          topicHeader = res.TopicHeader;
        }

        if (this.verbose) {
          logger.logToStderr(
            `Retrieving definition for web part ${args.options.webPartId ||
            args.options.standardWebPart}...`
          );
        }

        canvasContent = JSON.parse(res.CanvasContent1 || "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]");

        // Get the WebPart according to arguments
        return this.getWebPart(logger, args);
      })
      .then((webPart: ClientSideWebpart): Promise<void> => {
        if (this.verbose) {
          logger.logToStderr(`Setting client-side web part layout and properties...`);
        }

        this.setWebPartProperties(webPart, logger, args);

        // if no section exists (canvasContent array only has 1 default object), add a default section (1 col)
        if (canvasContent.length === 1) {
          const defaultSection: Control = {
            position: {
              controlIndex: 1,
              sectionIndex: 1,
              zoneIndex: 1,
              sectionFactor: 12,
              layoutIndex: 1,
            },
            emphasis: {},
            displayMode: 2
          };
          canvasContent.unshift(defaultSection);
        }

        // get unique zoneIndex values given each section can have 1 or more
        // columns each assigned to the zoneIndex of the corresponding section
        const zoneIndices: number[] = canvasContent
          .filter(c => c.position)
          .map(c => c.position.zoneIndex)
          .filter((value: number, index: number, array: number[]): boolean => {
            return array.indexOf(value) === index;
          })
          .sort((a, b) => a - b);

        // get section number. if not specified, get the last section
        const section: number = args.options.section || zoneIndices.length;
        if (section > zoneIndices.length) {
          return Promise.reject(`Invalid section '${section}'`);
        }

        // zoneIndex that represents the section where the web part should be added
        const zoneIndex: number = zoneIndices[section - 1];

        const column: number = args.options.column || 1;
        // we need the index of the control in the array so that we know which
        // item to replace or where to add the web part
        const controlIndex: number = canvasContent
          .findIndex(c => c.position &&
            c.position.zoneIndex === zoneIndex &&
            c.position.sectionIndex === column);
        if (controlIndex === -1) {
          return Promise.reject(`Invalid column '${args.options.column}'`);
        }

        // get the first control that matches section and column
        // if it's a empty column, it should be replaced with the web part
        // if it's a web part, then we need to determine if there are other
        // web parts and where in the array the new web part should be put
        const control: Control = canvasContent[controlIndex];
        const webPartControl: Control = this.extend({
          controlType: 3,
          displayMode: 2,
          id: webPart.id,
          position: Object.assign({}, control.position),
          webPartId: webPart.webPartId,
          emphasis: {}
        }, webPart);

        if (!control.controlType) {
          // it's an empty column so we need to replace it with the web part
          // ignore the specified order
          webPartControl.position.controlIndex = 1;
          canvasContent.splice(controlIndex, 1, webPartControl);
        }
        else {
          // it's a web part so we should find out where to put the web part in
          // the array of page controls

          // get web part index values to determine where to add the current
          // web part
          const controlIndices: number[] = canvasContent
            .filter(c => c.position &&
              c.position.zoneIndex === zoneIndex &&
              c.position.sectionIndex === column)
            .map(c => c.position.controlIndex as number)
            .sort((a, b) => a - b);

          // get the controlIndex of the web part before each the new web part
          // should be added
          if (!args.options.order ||
            args.options.order > controlIndices.length) {
            const controlIndex: number = controlIndices.pop() as number;
            const webPartIndex: number = canvasContent
              .findIndex(c => c.position &&
                c.position.zoneIndex === zoneIndex &&
                c.position.sectionIndex === column &&
                c.position.controlIndex === controlIndex);

            canvasContent.splice(webPartIndex + 1, 0, webPartControl);
          }
          else {
            const controlIndex: number = controlIndices[args.options.order - 1];
            const webPartIndex: number = canvasContent
              .findIndex(c => c.position &&
                c.position.zoneIndex === zoneIndex &&
                c.position.sectionIndex === column &&
                c.position.controlIndex === controlIndex);

            canvasContent.splice(webPartIndex, 0, webPartControl);
          }

          // reset order to ensure there are no gaps
          const webPartsInColumn: Control[] = canvasContent
            .filter(c => c.position &&
              c.position.zoneIndex === zoneIndex &&
              c.position.sectionIndex === column);
          let i: number = 1;
          webPartsInColumn.forEach(w => {
            w.position.controlIndex = i++;
          });
        }

        if (authorByline) {
          pageData.AuthorByline = authorByline;
        }
        if (bannerImageUrl) {
          pageData.BannerImageUrl = bannerImageUrl;
        }
        if (description) {
          pageData.Description = description;
        }
        if (title) {
          pageData.Title = title;
        }
        if (topicHeader) {
          pageData.TopicHeader = topicHeader;
        }
        if (layoutWebpartsContent) {
          pageData.LayoutWebpartsContent = layoutWebpartsContent;
        }
        if (canvasContent) {
          pageData.CanvasContent1 = JSON.stringify(canvasContent);
        }

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/sitepages/pages/GetByUrl('sitepages/${encodeURIComponent(pageFullName)}')/SavePageAsDraft`,
          headers: {
            'X-HTTP-Method': 'MERGE',
            'IF-MATCH': '*',
            'content-type': 'application/json;odata=nometadata',
            accept: 'application/json;odata=nometadata'
          },
          data: pageData,
          responseType: 'json'
        };

        return request.post(requestOptions);
      })
      .then(_ => cb())
      .catch((err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getWebPart(logger: Logger, args: CommandArgs): Promise<any> {
    return new Promise<any>((resolve: (webPart: any) => void, reject: (error: any) => void): void => {
      const standardWebPart: string | undefined = args.options.standardWebPart;

      const webPartId = standardWebPart
        ? StandardWebPartUtils.getWebPartId(standardWebPart as StandardWebPart)
        : args.options.webPartId;

      if (this.debug) {
        logger.logToStderr(`StandardWebPart: ${standardWebPart}`);
        logger.logToStderr(`WebPartId: ${webPartId}`);
      }

      const requestOptions: any = {
        url: `${args.options.webUrl}/_api/web/getclientsidewebparts()`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      request
        .get<{ value: ClientSidePageComponent[] }>(requestOptions)
        .then((res: { value: ClientSidePageComponent[] }): void => {
          const webPartDefinition = res.value.filter((c) => c.Id.toLowerCase() === (webPartId as string).toLowerCase() || c.Id.toLowerCase() === `{${(webPartId as string).toLowerCase()}}`);
          if (webPartDefinition.length === 0) {
            reject(new Error(`There is no available WebPart with Id ${webPartId}.`));
            return;
          }

          if (this.debug) {
            logger.logToStderr('WebPart definition:');
            logger.logToStderr(webPartDefinition);
            logger.logToStderr('');
          }

          if (this.verbose) {
            logger.logToStderr(`Creating instance from definition of WebPart ${webPartId}...`);
          }
          const component: ClientSidePageComponent = webPartDefinition[0];
          const id: string = v4();
          const componentId: string = component.Id.replace(/^\{|\}$/g, "").toLowerCase();
          const manifest: any = JSON.parse(component.Manifest);
          const preconfiguredEntries = manifest.preconfiguredEntries[0];
          const webPart = {
            id,
            webPartData: {
              dataVersion: "1.0",
              description: preconfiguredEntries.description.default,
              id: componentId,
              instanceId: id,
              properties: preconfiguredEntries.properties,
              title: preconfiguredEntries.title.default,
            },
            webPartId: componentId,
          };
          resolve(webPart);
        }, (error: any): void => {
          reject(error);
        });
    });
  }

  private setWebPartProperties(webPart: ClientSideWebpart, logger: Logger, args: CommandArgs): void {
    if (args.options.webPartProperties) {
      if (this.debug) {
        logger.logToStderr('WebPart properties: ');
        logger.logToStderr(args.options.webPartProperties);
        logger.logToStderr('');
      }

      try {
        const properties: any = JSON.parse(args.options.webPartProperties);
        (webPart as any).webPartData.properties = this.extend((webPart as any).webPartData.properties, properties)
      }
      catch {
      }
    }

    if (args.options.webPartData) {
      if (this.debug) {
        logger.logToStderr('WebPart data:');
        logger.logToStderr(args.options.webPartData);
        logger.logToStderr('');
      }

      const webPartData = JSON.parse(args.options.webPartData);
      (webPart as any).webPartData = this.extend((webPart as any).webPartData, webPartData);
      webPart.id = (webPart as any).webPartData.instanceId;
    }
  }

  /**
   * Provides functionality to extend the given object by doing a shallow copy
   *
   * @param target The object to which properties will be copied
   * @param source The source object from which properties will be copied
   *
   */
  private extend(target: any, source: any): any {
    return Object.getOwnPropertyNames(source)
      .reduce((t: any, v: string) => {
        t[v] = source[v];
        return t;
      }, target);
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-n, --pageName <pageName>'
      },
      {
        option: '--standardWebPart [standardWebPart]'
      },
      {
        option: '--webPartId [webPartId]'
      },
      {
        option: '--webPartProperties [webPartProperties]'
      },
      {
        option: '--webPartData [webPartData]'
      },
      {
        option: '--section [section]'
      },
      {
        option: '--column [column]'
      },
      {
        option: '--order [order]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.standardWebPart && !args.options.webPartId) {
      return 'Specify either the standardWebPart or the webPartId option';
    }

    if (args.options.standardWebPart && args.options.webPartId) {
      return 'Specify either the standardWebPart or the webPartId option but not both';
    }

    if (args.options.webPartId && !Utils.isValidGuid(args.options.webPartId)) {
      return `The webPartId '${args.options.webPartId}' is not a valid GUID`;
    }

    if (args.options.standardWebPart && !StandardWebPartUtils.isValidStandardWebPartType(args.options.standardWebPart)) {
      return `${args.options.standardWebPart} is not a valid standard web part type`;
    }

    if (args.options.webPartProperties && args.options.webPartData) {
      return 'Specify webPartProperties or webPartData but not both';
    }

    if (args.options.webPartProperties) {
      try {
        JSON.parse(args.options.webPartProperties);
      }
      catch (e) {
        return `Specified webPartProperties is not a valid JSON string. Input: ${args.options
          .webPartProperties}. Error: ${e}`;
      }
    }

    if (args.options.webPartData) {
      try {
        JSON.parse(args.options.webPartData);
      }
      catch (e) {
        return `Specified webPartData is not a valid JSON string. Input: ${args.options
          .webPartData}. Error: ${e}`;
      }
    }

    if (args.options.section && (!isNumber(args.options.section) || args.options.section < 1)) {
      return 'The value of parameter section must be 1 or higher';
    }

    if (args.options.column && (!isNumber(args.options.column) || args.options.column < 1)) {
      return 'The value of parameter column must be 1 or higher';
    }

    return SpoCommand.isValidSharePointUrl(args.options.webUrl);
  }
}

module.exports = new SpoPageClientSideWebPartAddCommand();
