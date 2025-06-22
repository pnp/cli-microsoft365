import { v4 } from 'uuid';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { BackgroundControl, Control } from './canvasContent.js';
import { CanvasColumnFactorType, CanvasSectionTemplate } from './clientsidepages.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  pageName: string;
  webUrl: string;
  sectionTemplate: string;
  order?: number;
  zoneEmphasis?: string;
  isLayoutReflowOnTop?: boolean;
  isCollapsibleSection?: boolean;
  showDivider?: boolean;
  iconAlignment?: string;
  isExpanded?: boolean;
  gradientText?: string;
  imageUrl?: string;
  imageHeight?: number;
  imageWidth?: number;
  fillMode?: string;
  useLightText?: boolean;
  overlayColor?: string;
  overlayOpacity?: number;
  collapsibleTitle?: string;
  zoneReflowStrategy?: string;
  zoneHeight?: number;
  headingLevel?: number;
}

class SpoPageSectionAddCommand extends SpoCommand {
  public readonly sectionTemplate: string[] = ['OneColumn', 'OneColumnFullWidth', 'TwoColumn', 'ThreeColumn', 'TwoColumnLeft', 'TwoColumnRight', 'Vertical', 'Flexible'];
  public readonly zoneEmphasis: string[] = ['None', 'Neutral', 'Soft', 'Strong', 'Image', 'Gradient'];
  public readonly iconAlignment: string[] = ['Left', 'Right'];
  public readonly fillMode: string[] = ['ScaleToFill', 'ScaleToFit', 'Tile', 'OriginalSize'];
  public readonly zoneReflowStrategy: string[] = ['TopToBottom', 'LeftToRight'];
  readonly MINIMUM_ZONE_HEIGHT: number = 34;

  public get name(): string {
    return commands.PAGE_SECTION_ADD;
  }

  public get description(): string {
    return 'Adds section to modern page';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        order: typeof args.options.order !== 'undefined',
        zoneEmphasis: typeof args.options.zoneEmphasis !== 'undefined',
        isLayoutReflowOnTop: !!args.options.isLayoutReflowOnTop,
        isCollapsibleSection: !!args.options.isCollapsibleSection,
        showDivider: !!typeof args.options.showDivider,
        iconAlignment: typeof args.options.iconAlignment !== 'undefined',
        isExpanded: !!args.options.isExpanded,
        gradientText: typeof args.options.gradientText !== 'undefined',
        imageUrl: typeof args.options.imageUrl !== 'undefined',
        imageHeight: typeof args.options.imageHeight !== 'undefined',
        imageWidth: typeof args.options.imageWidth !== 'undefined',
        fillMode: typeof args.options.fillMode !== 'undefined',
        useLightText: !!args.options.useLightText,
        overlayColor: typeof args.options.overlayColor !== 'undefined',
        overlayOpacity: typeof args.options.overlayOpacity !== 'undefined',
        collapsibleTitle: typeof args.options.collapsibleTitle !== 'undefined',
        zoneReflowStrategy: typeof args.options.zoneReflowStrategy !== 'undefined',
        zoneHeight: typeof args.options.zoneHeight !== 'undefined',
        headingLevel: typeof args.options.headingLevel !== 'undefined'
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
        option: '-t, --sectionTemplate <sectionTemplate>',
        autocomplete: this.sectionTemplate
      },
      {
        option: '--order [order]'
      },
      {
        option: '--zoneEmphasis [zoneEmphasis]',
        autocomplete: this.zoneEmphasis
      },
      {
        option: '--isLayoutReflowOnTop'
      },
      {
        option: '--isCollapsibleSection'
      },
      {
        option: '--showDivider'
      },
      {
        option: '--iconAlignment [iconAlignment]',
        autocomplete: this.iconAlignment
      },
      {
        option: '--isExpanded'
      },
      {
        option: '--gradientText [gradientText]'
      },
      {
        option: '--imageUrl [imageUrl]'
      },
      {
        option: '--imageHeight [imageHeight]'
      },
      {
        option: '--imageWidth [imageWidth]'
      },
      {
        option: '--fillMode [fillMode]',
        autocomplete: this.fillMode
      },
      {
        option: '--useLightText'
      },
      {
        option: '--overlayColor [overlayColor]'
      },
      {
        option: '--overlayOpacity [overlayOpacity]'
      },
      {
        option: '--collapsibleTitle [collapsibleTitle]'
      },
      {
        option: '--zoneReflowStrategy [zoneReflowStrategy]',
        autocomplete: this.zoneReflowStrategy
      },
      {
        option: '--zoneHeight [zoneHeight]'
      },
      {
        option: '--headingLevel [headingLevel]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!(args.options.sectionTemplate in CanvasSectionTemplate)) {
          return `${args.options.sectionTemplate} is not a valid section template. Allowed values are OneColumn|OneColumnFullWidth|TwoColumn|ThreeColumn|TwoColumnLeft|TwoColumnRight|Vertical|Flexible`;
        }

        if (typeof args.options.order !== 'undefined') {
          if (!Number.isInteger(args.options.order) || args.options.order < 1) {
            return 'The value of parameter order must be 1 or higher';
          }
        }

        if (typeof args.options.zoneEmphasis !== 'undefined') {
          if (!this.zoneEmphasis.some(zoneEmphasisValue => zoneEmphasisValue.toLowerCase() === args.options.zoneEmphasis?.toLowerCase())) {
            return `The value of parameter zoneEmphasis must be ${this.zoneEmphasis.join(', ')}`;
          }
        }

        if (typeof args.options.isLayoutReflowOnTop !== 'undefined') {
          if (args.options.sectionTemplate !== 'Vertical') {
            return 'Specify isLayoutReflowOnTop when the sectionTemplate is set to Vertical.';
          }
        }

        if (typeof args.options.iconAlignment !== 'undefined') {
          if (!this.iconAlignment.some(iconAlignmentValue => iconAlignmentValue.toLowerCase() === args.options.iconAlignment?.toLowerCase())) {
            return `The value of parameter iconAlignment must be ${this.iconAlignment.join(', ')}`;
          }
        }

        if (typeof args.options.fillMode !== 'undefined') {
          if (!this.fillMode.some(fillModeValue => fillModeValue.toLowerCase() === args.options.fillMode?.toLowerCase())) {
            return `The value of parameter fillMode must be ${this.fillMode.join(', ')}`;
          }
        }

        if (typeof args.options.zoneReflowStrategy !== 'undefined') {
          if (!this.zoneReflowStrategy.some(zoneReflowStrategyValue => zoneReflowStrategyValue.toLowerCase() === args.options.zoneReflowStrategy?.toLowerCase())) {
            return `The value of parameter zoneReflowStrategy must be ${this.zoneReflowStrategy.join(', ')}`;
          }
        }

        if (typeof args.options.headingLevel !== 'undefined') {
          if (![2, 3, 4].some(headingLevelValue => headingLevelValue === args.options.headingLevel)) {
            return `The value of parameter headingLevel must be 2, 3 or 4`;
          }
        }

        if (args.options.zoneEmphasis?.toLowerCase() !== 'image' && (args.options.imageUrl || args.options.imageWidth ||
          args.options.imageHeight || args.options.fillMode)) {
          return 'Specify imageUrl, imageWidth, imageHeight or fillMode only when zoneEmphasis is set to Image';
        }

        if (args.options.zoneEmphasis?.toLowerCase() === 'image' && !args.options.imageUrl) {
          return 'Specify imageUrl when zoneEmphasis is set to Image';
        }

        if (args.options.zoneEmphasis?.toLowerCase() !== 'gradient' && args.options.gradientText) {
          return 'Specify gradientText only when zoneEmphasis is set to Gradient';
        }

        if (args.options.zoneEmphasis?.toLowerCase() === 'gradient' && !args.options.gradientText) {
          return 'Specify gradientText when zoneEmphasis is set to Gradient';
        }

        if (args.options.overlayOpacity && (args.options.overlayOpacity < 0 || args.options.overlayOpacity > 100)) {
          return 'The value of parameter overlayOpacity must be between 0 and 100';
        }

        if (args.options.overlayColor && !/^#[0-9a-f]{6}$/i.test(args.options.overlayColor)) {
          return 'The value of parameter overlayColor must be a valid hex color';
        }

        if (!(args.options.zoneEmphasis && ['image', 'gradient'].includes(args.options.zoneEmphasis.toLowerCase())) && (args.options.overlayColor || args.options.overlayOpacity || args.options.useLightText)) {
          return 'Specify overlayColor or overlayOpacity only when zoneEmphasis is set to Image or Gradient';
        }

        if (args.options.sectionTemplate?.toLowerCase() !== 'flexible' && (args.options.zoneReflowStrategy || args.options.zoneHeight)) {
          return 'Specify zoneReflowStrategy or zoneHeight only when sectionTemplate is set to Flexible';
        }

        if (typeof args.options.zoneHeight !== 'undefined') {
          if (!Number.isInteger(args.options.zoneHeight) || args.options.zoneHeight < this.MINIMUM_ZONE_HEIGHT) {
            return `The value of parameter zoneHeight must be ${this.MINIMUM_ZONE_HEIGHT} or higher`;
          }
        }

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  #initTypes(): void {
    this.types.string = ['pageName', 'webUrl', 'sectionTemplate', 'zoneEmphasis', 'iconAlignment', 'gradientText', 'imageUrl', 'fillMode', 'overlayColor', 'collapsibleTitle', 'zoneReflowStrategy'];
    this.types.boolean = ['isLayoutReflowOnTop', 'isCollapsibleSection', 'showDivider', 'isExpanded', 'useLightText'];
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
      let requestOptions: any = {
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
        requestOptions = {
          url: `${args.options.webUrl}/_api/sitepages/pages/GetByUrl('sitepages/${formatting.encodeQueryParameter(pageFullName)}')/checkoutpage`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        await request.post(requestOptions);
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
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 0, args, zoneId));
        break;
      case 'TwoColumn':
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 6, args, zoneId));
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 6, args, zoneId));
        break;
      case 'ThreeColumn':
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 4, args, zoneId));
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 4, args, zoneId));
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 4, args, zoneId));
        break;
      case 'TwoColumnLeft':
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 8, args, zoneId));
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 4, args, zoneId));
        break;
      case 'TwoColumnRight':
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 4, args, zoneId));
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 8, args, zoneId));
        break;
      case 'Vertical':
        columns.push(this.getVerticalColumn(args, zoneId));
        break;
      case 'Flexible':
        columns.push(this.getFlexibleColumn(zoneIndex, sectionIndex++, args, zoneId));
        break;
      case 'OneColumn':
      default:
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 12, args, zoneId));
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
    columnValue.zoneReflowStrategy = { axis: args.options.zoneReflowStrategy ? this.zoneReflowStrategy.indexOf(args.options.zoneReflowStrategy) : 0 };
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
      backgroundDetails.fillMode = fillMode ? this.fillMode.indexOf(fillMode) : 0;
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