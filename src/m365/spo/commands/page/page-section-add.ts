import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { Control } from './canvasContent.js';
import { CanvasSectionTemplate, ZoneEmphasis } from './clientsidepages.js';

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
}

class SpoPageSectionAddCommand extends SpoCommand {
  public static readonly SectionTemplate: string[] = ['OneColumn', 'OneColumnFullWidth', 'TwoColumn', 'ThreeColumn', 'TwoColumnLeft', 'TwoColumnRight', 'Vertical'];
  public static readonly ZoneEmphasis: string[] = ['None', 'Neutral', 'Soft', 'Strong'];

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
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        order: typeof args.options.order !== 'undefined',
        zoneEmphasis: typeof args.options.zoneEmphasis !== 'undefined',
        isLayoutReflowOnTop: !!args.options.isLayoutReflowOnTop
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
        autocomplete: SpoPageSectionAddCommand.SectionTemplate
      },
      {
        option: '--order [order]'
      },
      {
        option: '--zoneEmphasis [zoneEmphasis]',
        autocomplete: SpoPageSectionAddCommand.ZoneEmphasis
      },
      {
        option: '--isLayoutReflowOnTop'
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
          if (!(args.options.zoneEmphasis in ZoneEmphasis)) {
            return 'The value of parameter zoneEmphasis must be None|Neutral|Soft|Strong';
          }
        }

        if (typeof args.options.isLayoutReflowOnTop !== 'undefined') {
          if (args.options.sectionTemplate !== 'Vertical') {
            return 'Specify isLayoutReflowOnTop when the sectionTemplate is set to Vertical.';
          }
        }

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let pageFullName: string = args.options.pageName.toLowerCase();
    if (pageFullName.indexOf('.aspx') < 0) {
      pageFullName += '.aspx';
    }
    let canvasContent: Control[];

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
      const columns: Control[] = canvasContent
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
      // get the list of columns to insert based on the selected template
      const columnsToAdd: Control[] = this.getColumns(zoneIndex, args.options.sectionTemplate, args.options.zoneEmphasis, args.options.isLayoutReflowOnTop);
      // insert the column in the right place in the array so that
      // it stays sorted ascending by zoneIndex
      let pos: number = canvasContent.findIndex(c => typeof c.controlType === 'undefined' && c.position.zoneIndex > zoneIndex);
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

  private getColumns(zoneIndex: number, sectionTemplate: string, zoneEmphasis?: string, isLayoutReflowOnTop?: boolean): Control[] {
    const columns: Control[] = [];
    let sectionIndex: number = 1;

    switch (sectionTemplate) {
      case 'OneColumnFullWidth':
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 0, zoneEmphasis));
        break;
      case 'TwoColumn':
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 6, zoneEmphasis));
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 6, zoneEmphasis));
        break;
      case 'ThreeColumn':
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 4, zoneEmphasis));
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 4, zoneEmphasis));
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 4, zoneEmphasis));
        break;
      case 'TwoColumnLeft':
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 8, zoneEmphasis));
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 4, zoneEmphasis));
        break;
      case 'TwoColumnRight':
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 4, zoneEmphasis));
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 8, zoneEmphasis));
        break;
      case 'Vertical':
        columns.push(this.getVerticalColumn(zoneEmphasis, isLayoutReflowOnTop));
        break;
      case 'OneColumn':
      default:
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 12, zoneEmphasis));
        break;
    }

    return columns;
  }

  private getColumn(zoneIndex: number, sectionIndex: number, sectionFactor: number, zoneEmphasis?: string): Control {
    const columnValue: Control = {
      displayMode: 2,
      position: {
        zoneIndex: zoneIndex,
        sectionIndex: sectionIndex,
        sectionFactor: sectionFactor,
        layoutIndex: 1
      },
      emphasis: {
      }
    };

    if (zoneEmphasis) {
      const zoneEmphasisValue: number = ZoneEmphasis[zoneEmphasis as keyof typeof ZoneEmphasis];
      columnValue.emphasis = { zoneEmphasis: zoneEmphasisValue };
    }

    return columnValue;
  }

  private getVerticalColumn(zoneEmphasis?: string, isLayoutReflowOnTop?: boolean): Control {
    const columnValue: Control = this.getColumn(1, 1, 12, zoneEmphasis);
    columnValue.position.isLayoutReflowOnTop = isLayoutReflowOnTop !== undefined ? true : false;
    columnValue.position.layoutIndex = 2;
    columnValue.position.controlIndex = 1;

    return columnValue;
  }
}

export default new SpoPageSectionAddCommand();