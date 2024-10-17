import { v4 } from 'uuid';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { BackgroundControl, Control } from './canvasContent.js';
import { CanvasSectionTemplate } from './clientsidepages.js';

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
}

class SpoPageSectionAddCommand extends SpoCommand {
  public readonly sectionTemplate: string[] = ['OneColumn', 'OneColumnFullWidth', 'TwoColumn', 'ThreeColumn', 'TwoColumnLeft', 'TwoColumnRight', 'Vertical'];
  public readonly zoneEmphasis: string[] = ['None', 'Neutral', 'Soft', 'Strong', 'Image', 'Gradient'];
  public readonly iconAlignment: string[] = ['Left', 'Right'];
  public readonly fillMode: string[] = ['ScaleToFill', 'ScaleToFit', 'Tile', 'OriginalSize'];

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
        overlayOpacity: typeof args.options.overlayOpacity !== 'undefined'
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
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!(args.options.sectionTemplate in CanvasSectionTemplate)) {
          return `${args.options.sectionTemplate} is not a valid section template. Allowed values are OneColumn|OneColumnFullWidth|TwoColumn|ThreeColumn|TwoColumnLeft|TwoColumnRight|Vertical`;
        }

        if (typeof args.options.order !== 'undefined') {
          if (!Number.isInteger(args.options.order) || args.options.order < 1) {
            return 'The value of parameter order must be 1 or higher';
          }
        }

        if (typeof args.options.zoneEmphasis !== 'undefined') {
          if (!this.zoneEmphasis.some(zoneEmphasisValue => zoneEmphasisValue.toLocaleLowerCase() === args.options.zoneEmphasis?.toLowerCase())) {
            return `The value of parameter zoneEmphasis must be ${this.zoneEmphasis.join(', ')}`;
          }
        }

        if (typeof args.options.isLayoutReflowOnTop !== 'undefined') {
          if (args.options.sectionTemplate !== 'Vertical') {
            return 'Specify isLayoutReflowOnTop when the sectionTemplate is set to Vertical.';
          }
        }

        if (typeof args.options.iconAlignment !== 'undefined') {
          if (!this.iconAlignment.some(iconAlignmentValue => iconAlignmentValue.toLocaleLowerCase() === args.options.iconAlignment?.toLowerCase())) {
            return `The value of parameter iconAlignment must be ${this.iconAlignment.join(', ')}`;
          }
        }

        if (typeof args.options.fillMode !== 'undefined') {
          if (!this.fillMode.some(fillModeValue => fillModeValue.toLocaleLowerCase() === args.options.fillMode?.toLowerCase())) {
            return `The value of parameter fillMode must be ${this.fillMode.join(', ')}`;
          }
        }

        if (args.options.zoneEmphasis?.toLocaleLowerCase() !== 'image' && (args.options.imageUrl || args.options.imageWidth ||
          args.options.imageHeight || args.options.fillMode)) {
          return 'Specify imageUrl, imageWidth, imageHeight or fillMode only when zoneEmphasis is set to Image';
        }

        if (args.options.zoneEmphasis?.toLocaleLowerCase() === 'image' && !args.options.imageUrl) {
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

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  #initTypes(): void {
    this.types.string = ['pageName', 'webUrl', 'sectionTemplate', 'zoneEmphasis', 'iconAlignment', 'gradientText', 'imageUrl', 'fillMode', 'overlayColor'];
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

      // get columns
      const columns: (Control | BackgroundControl)[] = canvasContent
        .filter(c => typeof c.controlType === 'undefined');
      // get unique zoneIndex values given each section can have 1 or more
      // columns each assigned to the zoneIndex of the corresponding section
      const zoneIndices: number[] = columns
        .map(c => c.position.zoneIndex)
        .filter((value: number, index: number, array: number[]): boolean => {
          return array.indexOf(value) === index;
        })
        .sort();
      // zoneIndex for the new section to add
      const zoneIndex: number = this.getSectionIndex(zoneIndices, args.options.order);
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

      // get the list of columns to insert based on the selected template
      const columnsToAdd: Control[] = this.getColumns(zoneIndex, args, zoneId);
      // insert the column in the right place in the array so that
      // it stays sorted ascending by zoneIndex
      let pos: number = canvasContent.findIndex(c => typeof c.controlType === 'undefined' && c.position && c.position.zoneIndex > zoneIndex);
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

  private getSectionIndex(zoneIndices: number[], order?: number): number {
    // zoneIndex of the first column on the page
    const minIndex: number = zoneIndices.length === 0 ? 0 : zoneIndices[0];
    // zoneIndex of the last column on the page
    const maxIndex: number = zoneIndices.length === 0 ? 0 : zoneIndices[zoneIndices.length - 1];
    if (!order || order > zoneIndices.length) {
      // no order specified, add section to the end
      return maxIndex === 0 ? 1 : maxIndex * 2;
    }

    // add to the beginning
    if (order === 1) {
      return minIndex / 2;
    }

    return zoneIndices[order - 2] + ((zoneIndices[order - 1] - zoneIndices[order - 2]) / 2);
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
      case 'OneColumn':
      default:
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 12, args, zoneId));
        break;
    }

    return columns;
  }

  private getColumn(zoneIndex: number, sectionIndex: number, sectionFactor: number, args: CommandArgs, zoneId?: string): Control {
    const { zoneEmphasis, isCollapsibleSection, isExpanded, showDivider, iconAlignment } = args.options;
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

    if (zoneEmphasis && ['none', 'neutral', 'soft', 'strong'].includes(zoneEmphasis?.toLocaleLowerCase())) {
      // Just these zoneEmphasis values should be added to column emphasis
      const zoneEmphasisValue: number = ['none', 'neutral', 'soft', 'strong'].indexOf(zoneEmphasis.toLocaleLowerCase());
      columnValue.emphasis = { zoneEmphasis: zoneEmphasisValue };
    }

    if (isCollapsibleSection) {
      columnValue.zoneGroupMetadata = {
        type: 1,
        isExpanded: !!isExpanded,
        showDividerLine: !!showDivider,
        iconAlignment: iconAlignment && iconAlignment.toLocaleLowerCase() === "right" ? "right" : "left"
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
}

export default new SpoPageSectionAddCommand();