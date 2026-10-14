import { v4 } from 'uuid';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { StandardWebPart, StandardWebPartUtils } from '../../StandardWebPartTypes.js';
import { Control } from './canvasContent.js';
import { ClientSidePageProperties } from './ClientSidePageProperties.js';
import { ClientSidePageComponent, ClientSideWebpart } from './clientsidepages.js';
import { Page } from './Page.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  pageName: z.string().alias('n'),
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  standardWebPart: z.string()
    .refine(webPart => StandardWebPartUtils.isValidStandardWebPartType(webPart), {
      error: e => `${e.input} is not a valid standard web part type`
    })
    .optional(),
  webPartData: z.string()
    .refine(data => {
      try {
        JSON.parse(data);
        return true;
      }
      catch {
        return false;
      }
    }, {
      error: 'Specified webPartData is not a valid JSON string.'
    })
    .optional(),
  webPartId: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `The webPartId '${e.input}' is not a valid GUID`
    })
    .optional(),
  webPartProperties: z.string()
    .refine(properties => {
      try {
        JSON.parse(properties);
        return true;
      }
      catch {
        return false;
      }
    }, {
      error: 'Specified webPartProperties is not a valid JSON string.'
    })
    .optional(),
  section: z.number().int().min(1, { error: 'The value of parameter section must be 1 or higher' }).optional(),
  column: z.number().int().min(1, { error: 'The value of parameter column must be 1 or higher' }).optional(),
  order: z.number().optional(),
  verticalSection: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoPageClientSideWebPartAddCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_CLIENTSIDEWEBPART_ADD;
  }

  public get description(): string {
    return 'Adds a client-side web part to a modern page';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.standardWebPart, opts.webPartId].filter(o => o !== undefined).length === 1, {
        error: 'Specify one of the following options: standardWebPart, webPartId.',
        params: {
          customCode: 'optionSet',
          options: ['standardWebPart', 'webPartId']
        }
      })
      .refine(opts => !(opts.webPartProperties && opts.webPartData), {
        error: 'Specify webPartProperties or webPartData but not both'
      })
      .refine(opts => !(opts.section && opts.verticalSection), {
        error: 'Specify section or verticalSection but not both'
      })
      .refine(opts => !(opts.column && opts.verticalSection), {
        error: 'Use column in combination with section, not with verticalSection'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let canvasContent: Control[];

    let layoutWebpartsContent: string = "";
    let authorByline: string[] = [""];
    let bannerImageUrl: string = "";
    let description: string = "";
    let title: string = "";
    let topicHeader: string = "";
    const pageData: any = {};

    let pageFullName: string = args.options.pageName;
    if (args.options.pageName.indexOf('.aspx') < 0) {
      pageFullName += '.aspx';
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving page information...`);
    }

    try {
      let requestOptions: any = {
        url: `${args.options.webUrl}/_api/sitepages/pages/GetByUrl('sitepages/${formatting.encodeQueryParameter(pageFullName)}')`,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      let page = await request.get<ClientSidePageProperties>(requestOptions);
      if (!page.IsPageCheckedOutToCurrentUser) {
        page = await Page.checkout(pageFullName, args.options.webUrl, logger, this.verbose);
      }

      if (page) {
        layoutWebpartsContent = page.LayoutWebpartsContent;
        authorByline = page.AuthorByline;
        bannerImageUrl = page.BannerImageUrl;
        description = page.Description;
        title = page.Title;
        topicHeader = page.TopicHeader;
      }

      if (this.verbose) {
        await logger.logToStderr(
          `Retrieving definition for web part ${args.options.webPartId ||
          args.options.standardWebPart}...`
        );
      }

      canvasContent = JSON.parse(page.CanvasContent1 || "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]");

      // Get the WebPart according to arguments
      const webPart = await this.getWebPart(logger, args);
      if (this.verbose) {
        await logger.logToStderr(`Setting client-side web part layout and properties...`);
      }

      await this.setWebPartProperties(webPart, logger, args);

      const control: Control = this.getCorrectControl(canvasContent, args);
      const controlIndex: number = canvasContent.indexOf(control);
      const zoneIndex: number = control.position.zoneIndex;
      const column: number = control.position.sectionIndex;

      const webPartControl: Control = this.extend({
        controlType: 3,
        displayMode: 2,
        id: webPart.id,
        position: Object.assign({}, control.position),
        webPartId: webPart.webPartId,
        emphasis: {},
        zoneGroupMetadata: control.zoneGroupMetadata
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

      requestOptions = {
        url: `${args.options.webUrl}/_api/sitepages/pages/GetByUrl('sitepages/${formatting.encodeQueryParameter(pageFullName)}')/SavePageAsDraft`,
        headers: {
          'X-HTTP-Method': 'MERGE',
          'IF-MATCH': '*',
          'content-type': 'application/json;odata=nometadata',
          accept: 'application/json;odata=nometadata'
        },
        data: pageData,
        responseType: 'json'
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getWebPart(logger: Logger, args: CommandArgs): Promise<any> {
    const standardWebPart: string | undefined = args.options.standardWebPart;

    const webPartId = standardWebPart
      ? StandardWebPartUtils.getWebPartId(standardWebPart as StandardWebPart)
      : args.options.webPartId;

    if (this.debug) {
      await logger.logToStderr(`StandardWebPart: ${standardWebPart}`);
      await logger.logToStderr(`WebPartId: ${webPartId}`);
    }

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web/getclientsidewebparts()`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const response: { value: ClientSidePageComponent[] } = await request.get<{ value: ClientSidePageComponent[] }>(requestOptions);
    const webPartDefinition = response.value.filter((c) => c.Id.toLowerCase() === (webPartId as string).toLowerCase() || c.Id.toLowerCase() === `{${(webPartId as string).toLowerCase()}}`);
    if (webPartDefinition.length === 0) {
      throw new Error(`There is no available WebPart with Id ${webPartId}.`);
    }

    if (this.debug) {
      await logger.logToStderr('WebPart definition:');
      await logger.logToStderr(webPartDefinition);
      await logger.logToStderr('');
    }

    if (this.verbose) {
      await logger.logToStderr(`Creating instance from definition of WebPart ${webPartId}...`);
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
        title: preconfiguredEntries.title.default
      },
      webPartId: componentId
    };
    return webPart;
  }

  private async setWebPartProperties(webPart: ClientSideWebpart, logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.webPartProperties) {
      if (this.debug) {
        await logger.logToStderr('WebPart properties: ');
        await logger.logToStderr(args.options.webPartProperties);
        await logger.logToStderr('');
      }

      try {
        const properties: any = JSON.parse(args.options.webPartProperties);
        (webPart as any).webPartData.properties = this.extend((webPart as any).webPartData.properties, properties);
      }
      catch {
        // Do nothing
      }
    }

    if (args.options.webPartData) {
      if (this.debug) {
        await logger.logToStderr('WebPart data:');
        await logger.logToStderr(args.options.webPartData);
        await logger.logToStderr('');
      }

      const webPartData = JSON.parse(args.options.webPartData);
      (webPart as any).webPartData = this.extend((webPart as any).webPartData, webPartData);
      webPart.id = (webPart as any).webPartData.instanceId;
    }
  }

  private getCorrectControl(canvasContent: Control[], args: CommandArgs): Control {
    // get Vertical section 
    if (args.options.verticalSection) {
      let verticalSection = canvasContent
        .find(c => c.position?.layoutIndex === 2);

      //if vertical section does not exist, create it
      if (!verticalSection) {
        verticalSection = {
          position: {
            controlIndex: 1,
            sectionIndex: 1,
            zoneIndex: 1,
            sectionFactor: 12,
            layoutIndex: 2
          },
          emphasis: {},
          displayMode: 2
        };
        canvasContent.unshift(verticalSection);
      }
      return verticalSection;
    }

    // if no section exists (canvasContent array only has 1 default object), add a default section (1 col)
    if (canvasContent.length === 1) {
      const defaultSection: Control = {
        position: {
          controlIndex: 1,
          sectionIndex: 1,
          zoneIndex: 1,
          sectionFactor: 12,
          layoutIndex: 1
        },
        emphasis: {},
        displayMode: 2
      };
      canvasContent.unshift(defaultSection);
    }

    // get unique zoneIndex values given each section can have 1 or more
    // columns each assigned to the zoneIndex of the corresponding section
    const zoneIndices: number[] = canvasContent
      // Exclude the vertical section
      .filter(c => c.position)
      .map(c => c.position.zoneIndex)
      .filter((value: number, index: number, array: number[]): boolean => {
        return array.indexOf(value) === index;
      })
      .sort((a, b) => a - b);

    // get section number. if not specified, get the last section
    const section: number = args.options.section || zoneIndices.length;
    if (section > zoneIndices.length) {
      throw `Invalid section '${section}'`;
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
      throw `Invalid column '${args.options.column}'`;
    }

    // get the first control that matches section and column
    // if it's a empty column, it should be replaced with the web part
    // if it's a web part, then we need to determine if there are other
    // web parts and where in the array the new web part should be put
    return canvasContent[controlIndex];
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
}

export default new SpoPageClientSideWebPartAddCommand();
