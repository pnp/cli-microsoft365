import { v4 } from 'uuid';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { BackgroundControl, Control } from './canvasContent.js';
import { CanvasColumnFactorType } from './clientsidepages.js';
import { Page } from './Page.js';

const sectionTemplates = ['OneColumn', 'OneColumnFullWidth', 'TwoColumn', 'ThreeColumn', 'TwoColumnLeft', 'TwoColumnRight', 'Vertical', 'Flexible'] as const;
const zoneEmphasisValues = ['None', 'Neutral', 'Soft', 'Strong', 'Image', 'Gradient'] as const;
const iconAlignmentValues = ['Left', 'Right'] as const;
const fillModeValues = ['ScaleToFill', 'ScaleToFit', 'Tile', 'OriginalSize'] as const;
const zoneReflowStrategyValues = ['TopToBottom', 'LeftToRight'] as const;
const MINIMUM_ZONE_HEIGHT = 34;

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  pageName: z.string().alias('n'),
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  sectionTemplate: z.enum(sectionTemplates).alias('t'),
  order: z.number().int().min(1, { error: 'The value of parameter order must be 1 or higher' }).optional(),
  zoneEmphasis: z.preprocess(val => {
    const target = String(val).toLowerCase();
    return zoneEmphasisValues.find(v => v.toLowerCase() === target) ?? val;
  }, z.enum(zoneEmphasisValues)).optional(),
  isLayoutReflowOnTop: z.boolean().optional(),
  isCollapsibleSection: z.boolean().optional(),
  showDivider: z.boolean().optional(),
  iconAlignment: z.preprocess(val => {
    const target = String(val).toLowerCase();
    return iconAlignmentValues.find(v => v.toLowerCase() === target) ?? val;
  }, z.enum(iconAlignmentValues)).optional(),
  isExpanded: z.boolean().optional(),
  gradientText: z.string().optional(),
  imageUrl: z.string().optional(),
  imageHeight: z.number().optional(),
  imageWidth: z.number().optional(),
  fillMode: z.preprocess(val => {
    const target = String(val).toLowerCase();
    return fillModeValues.find(v => v.toLowerCase() === target) ?? val;
  }, z.enum(fillModeValues)).optional(),
  useLightText: z.boolean().optional(),
  overlayColor: z.string().optional(),
  overlayOpacity: z.number().optional(),
  collapsibleTitle: z.string().optional(),
  zoneReflowStrategy: z.preprocess(val => {
    const target = String(val).toLowerCase();
    return zoneReflowStrategyValues.find(v => v.toLowerCase() === target) ?? val;
  }, z.enum(zoneReflowStrategyValues)).optional(),
  zoneHeight: z.number().int().min(MINIMUM_ZONE_HEIGHT, { error: `The value of parameter zoneHeight must be ${MINIMUM_ZONE_HEIGHT} or higher` }).optional(),
  headingLevel: z.number()
    .refine(level => [2, 3, 4].includes(level), {
      error: 'The value of parameter headingLevel must be 2, 3 or 4'
    })
    .optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoPageSectionAddCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_SECTION_ADD;
  }

  public get description(): string {
    return 'Adds section to modern page';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => opts.isLayoutReflowOnTop === undefined || opts.sectionTemplate === 'Vertical', {
        error: 'Specify isLayoutReflowOnTop when the sectionTemplate is set to Vertical.'
      })
      .refine(opts => opts.zoneEmphasis === 'Image' || !(opts.imageUrl || opts.imageWidth || opts.imageHeight || opts.fillMode), {
        error: 'Specify imageUrl, imageWidth, imageHeight or fillMode only when zoneEmphasis is set to Image'
      })
      .refine(opts => opts.zoneEmphasis !== 'Image' || !!opts.imageUrl, {
        error: 'Specify imageUrl when zoneEmphasis is set to Image'
      })
      .refine(opts => opts.zoneEmphasis === 'Gradient' || !opts.gradientText, {
        error: 'Specify gradientText only when zoneEmphasis is set to Gradient'
      })
      .refine(opts => opts.zoneEmphasis !== 'Gradient' || !!opts.gradientText, {
        error: 'Specify gradientText when zoneEmphasis is set to Gradient'
      })
      .refine(opts => opts.overlayOpacity === undefined || (opts.overlayOpacity >= 0 && opts.overlayOpacity <= 100), {
        error: 'The value of parameter overlayOpacity must be between 0 and 100'
      })
      .refine(opts => !opts.overlayColor || /^#[0-9a-f]{6}$/i.test(opts.overlayColor), {
        error: 'The value of parameter overlayColor must be a valid hex color'
      })
      .refine(opts => ['Image', 'Gradient'].includes(opts.zoneEmphasis ?? '') || !(opts.overlayColor || opts.overlayOpacity || opts.useLightText), {
        error: 'Specify overlayColor or overlayOpacity only when zoneEmphasis is set to Image or Gradient'
      })
      .refine(opts => opts.sectionTemplate === 'Flexible' || !(opts.zoneReflowStrategy || opts.zoneHeight), {
        error: 'Specify zoneReflowStrategy or zoneHeight only when sectionTemplate is set to Flexible'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let pageFullName: string = args.options.pageName.toLowerCase();
    if (!pageFullName.endsWith('.aspx')) {
      pageFullName += '.aspx';
    }

    let canvasContent: (Control | BackgroundControl)[];

    if (this.verbose) {
      await logger.logToStderr(`Retrieving page information...`);
    }

    try {
      let requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/sitepages/pages/GetByUrl('sitepages/${formatting.encodeQueryParameter(pageFullName)}')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const res = await request.get<{ CanvasContent1: string; IsPageCheckedOutToCurrentUser: boolean }>(requestOptions);
      canvasContent = JSON.parse(res.CanvasContent1 || "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]");

      if (args.options.sectionTemplate === 'OneColumnFullWidth') {
        this.ensureFullWidthSectionCanBeAdded(canvasContent);
      }

      if (args.options.sectionTemplate === 'Vertical') {
        this.ensureVerticalSectionCanBeAdded(canvasContent);
      }

      if (!res.IsPageCheckedOutToCurrentUser) {
        await Page.checkout(pageFullName, args.options.webUrl, logger, this.verbose);
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

      // Add a new zoneIndex  at the end of the array
      zoneIndices.push(zoneIndices.length > 0 ? zoneIndices[zoneIndices.length - 1] + 1 : 1);

      // get section number. if not specified, get the last section
      let section: number = args.options.order || zoneIndices.length;
      if (section > zoneIndices.length) {
        section = zoneIndices.length;
      }

      // zoneIndex that represents the section where the web part should be added
      const zoneIndex: number = zoneIndices[section - 1];
      let zoneId: string | undefined;

      let backgroundControlToAdd: BackgroundControl | undefined = undefined;

      if (args.options.zoneEmphasis && ['image', 'gradient'].includes(args.options.zoneEmphasis.toLowerCase())) {
        zoneId = v4();

        // get background control based on control type 14
        const backgroundControl = canvasContent.find(c => c.controlType === 14) as BackgroundControl;
        backgroundControlToAdd = this.setBackgroundControl(zoneId, backgroundControl, args);

        if (!backgroundControl) {
          canvasContent.push(backgroundControlToAdd);
        }
      }

      // Increment the zoneIndex of all columns that are greater than or equal to the new zoneIndex
      canvasContent.forEach((c: Control | BackgroundControl) => {
        if (c.position && c.position.zoneIndex >= zoneIndex) {
          c.position.zoneIndex += 1;
        }
      });

      // get the list of columns to insert based on the selected template
      const columnsToAdd: Control[] = this.getColumns(zoneIndex, args, zoneId);
      // insert the column in the right place in the array so that
      // it stays sorted ascending by zoneIndex
      let pos: number = canvasContent.findIndex(c => c.position && c.position.zoneIndex >= zoneIndex);
      if (pos === -1) {
        pos = canvasContent.length - 1;
      }
      canvasContent.splice(pos, 0, ...columnsToAdd);

      requestOptions = {
        url: `${args.options.webUrl}/_api/sitepages/pages/GetByUrl('sitepages/${formatting.encodeQueryParameter(pageFullName)}')/savepage`,
        headers: {
          'accept': 'application/json;odata=nometadata',
          'content-type': 'application/json;odata=nometadata'
        },
        data: {
          CanvasContent1: JSON.stringify(canvasContent)
        },
        responseType: 'json'
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getColumns(zoneIndex: number, args: CommandArgs, zoneId?: string): Control[] {
    const columns: Control[] = [];
    let sectionIndex: number = 1;

    switch (args.options.sectionTemplate) {
      case 'OneColumnFullWidth':
        columns.push(this.getColumn(zoneIndex, sectionIndex, 0, args, zoneId));
        break;
      case 'TwoColumn':
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 6, args, zoneId));
        columns.push(this.getColumn(zoneIndex, sectionIndex, 6, args, zoneId));
        break;
      case 'ThreeColumn':
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 4, args, zoneId));
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 4, args, zoneId));
        columns.push(this.getColumn(zoneIndex, sectionIndex, 4, args, zoneId));
        break;
      case 'TwoColumnLeft':
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 8, args, zoneId));
        columns.push(this.getColumn(zoneIndex, sectionIndex, 4, args, zoneId));
        break;
      case 'TwoColumnRight':
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 4, args, zoneId));
        columns.push(this.getColumn(zoneIndex, sectionIndex, 8, args, zoneId));
        break;
      case 'Vertical':
        columns.push(this.getVerticalColumn(args, zoneId));
        break;
      case 'Flexible':
        columns.push(this.getFlexibleColumn(zoneIndex, sectionIndex, args, zoneId));
        break;
      case 'OneColumn':
      default:
        columns.push(this.getColumn(zoneIndex, sectionIndex, 12, args, zoneId));
        break;
    }

    return columns;
  }

  private getColumn(zoneIndex: number, sectionIndex: number, sectionFactor: CanvasColumnFactorType, args: CommandArgs, zoneId?: string): Control {
    const { zoneEmphasis, isCollapsibleSection, isExpanded, showDivider, iconAlignment, collapsibleTitle } = args.options;
    const columnValue: Control = {
      displayMode: 2,
      position: {
        zoneIndex: zoneIndex,
        sectionIndex: sectionIndex,
        sectionFactor: sectionFactor,
        layoutIndex: 1,
        zoneId: zoneId
      },
      emphasis: {
      }
    };

    if (zoneEmphasis && ['none', 'neutral', 'soft', 'strong'].includes(zoneEmphasis?.toLowerCase())) {
      // Just these zoneEmphasis values should be added to column emphasis
      const zoneEmphasisValue: number = ['none', 'neutral', 'soft', 'strong'].indexOf(zoneEmphasis.toLowerCase());
      columnValue.emphasis = { zoneEmphasis: zoneEmphasisValue };
    }

    if (isCollapsibleSection) {
      columnValue.zoneGroupMetadata = {
        type: 1,
        isExpanded: !!isExpanded,
        showDividerLine: !!showDivider,
        iconAlignment: iconAlignment && iconAlignment.toLowerCase() === "right" ? "right" : "left",
        displayName: collapsibleTitle,
        headingLevel: args.options.headingLevel ? args.options.headingLevel : 2 //2 is a default heading level
      };
    }

    return columnValue;
  }

  private getVerticalColumn(args: CommandArgs, zoneId?: string): Control {
    const columnValue: Control = this.getColumn(1, 1, 12, args, zoneId);
    columnValue.position.isLayoutReflowOnTop = args.options.isLayoutReflowOnTop !== undefined;
    columnValue.position.layoutIndex = 2;
    columnValue.position.controlIndex = 1;

    return columnValue;
  }

  private getFlexibleColumn(zoneIndex: number, sectionIndex: number, args: CommandArgs, zoneId?: string): Control {
    const columnValue: Control = this.getColumn(zoneIndex, sectionIndex, 100, args, zoneId);
    columnValue.zoneReflowStrategy = { axis: args.options.zoneReflowStrategy ? zoneReflowStrategyValues.indexOf(args.options.zoneReflowStrategy) : 0 };
    if (args.options.zoneHeight) {
      columnValue.zoneHeight = args.options.zoneHeight;
    }

    return columnValue;
  }

  private setBackgroundControl(zoneId: string, backgroundControl: BackgroundControl, args: CommandArgs): BackgroundControl {
    const { overlayColor, overlayOpacity, useLightText, imageUrl } = args.options;
    const backgroundDetails = this.getBackgroundDetails(args);

    if (!backgroundControl) {
      backgroundControl = {
        controlType: 14,
        webPartData: {
          properties: {
            zoneBackground: {
            }
          },
          serverProcessedContent: {
            htmlStrings: {},
            searchablePlainTexts: {},
            imageSources: {},
            links: {}
          },
          dataVersion: "1.0"
        }
      };
    }

    backgroundControl.webPartData.properties.zoneBackground[zoneId] = {
      ...backgroundDetails,
      useLightText: !!useLightText,
      overlay: {
        color: overlayColor ? overlayColor : "#FFFFFF",
        opacity: overlayOpacity ? overlayOpacity : 60
      }
    };

    if (imageUrl && backgroundControl.webPartData.serverProcessedContent.imageSources) {
      backgroundControl.webPartData.serverProcessedContent.imageSources[`zoneBackground.${zoneId}.imageData.url`] = imageUrl;
    }
    return backgroundControl;
  }

  private getBackgroundDetails(args: CommandArgs): any {
    const { gradientText, imageUrl, imageHeight, imageWidth, fillMode } = args.options;
    const backgroundDetails: any = {};

    if (gradientText) {
      backgroundDetails.type = "gradient";
      backgroundDetails.gradient = gradientText;
    }

    if (imageUrl) {
      backgroundDetails.type = "image";
      backgroundDetails.imageData = {
        source: 2,
        fileName: "sectionbackground.jpg",
        height: imageHeight ? imageHeight : 955,
        width: imageWidth ? imageWidth : 555
      };
      backgroundDetails.fillMode = fillMode ? fillModeValues.indexOf(fillMode) : 0;
    }

    return backgroundDetails;
  }

  private ensureFullWidthSectionCanBeAdded(canvasContent: (Control | BackgroundControl)[]): void {
    const hasVerticalSection = canvasContent.some((c: Control | BackgroundControl) =>
      c.position?.layoutIndex === 2 && c.position.sectionFactor === 12
    );

    if (hasVerticalSection) {
      throw "A vertical section already exists on the page. A full-width section cannot be added to a page that already has a vertical section.";
    }
  }

  private ensureVerticalSectionCanBeAdded(canvasContent: (Control | BackgroundControl)[]): void {
    const hasFullWidthSection = canvasContent.some((c: Control | BackgroundControl) =>
      c.position?.layoutIndex === 1 && c.position.sectionFactor === 0
    );

    if (hasFullWidthSection) {
      throw "A full-width section already exists on the page. A vertical section cannot be added to a page that already has a full-width section.";
    }
  }
}

export default new SpoPageSectionAddCommand();