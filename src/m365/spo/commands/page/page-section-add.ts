import { isNumber } from 'util';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { Control } from './canvasContent';
import { CanvasSectionTemplate } from './clientsidepages';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  pageName: string;
  webUrl: string;
  sectionTemplate: string;
  order?: number;
}

class SpoPageSectionAddCommand extends SpoCommand {
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
        order: typeof args.options.order !== 'undefined'
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
        option: '-t, --sectionTemplate <sectionTemplate>'
      },
      {
        option: '--order [order]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!(args.options.sectionTemplate in CanvasSectionTemplate)) {
          return `${args.options.sectionTemplate} is not a valid section template. Allowed values are OneColumn|OneColumnFullWidth|TwoColumn|ThreeColumn|TwoColumnLeft|TwoColumnRight`;
        }

        if (typeof args.options.order !== 'undefined') {
          if (!isNumber(args.options.order) || args.options.order < 1) {
            return 'The value of parameter order must be 1 or higher';
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
      logger.logToStderr(`Retrieving page information...`);
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

        return request.post(requestOptions);
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
      const columnsToAdd: Control[] = this.getColumns(zoneIndex, args.options.sectionTemplate);
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

  private getColumns(zoneIndex: number, sectionTemplate: string): Control[] {
    const columns: Control[] = [];
    let sectionIndex: number = 1;

    switch (sectionTemplate) {
      case 'OneColumnFullWidth':
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 0));
        break;
      case 'TwoColumn':
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 6));
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 6));
        break;
      case 'ThreeColumn':
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 4));
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 4));
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 4));
        break;
      case 'TwoColumnLeft':
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 8));
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 4));
        break;
      case 'TwoColumnRight':
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 4));
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 8));
        break;
      case 'OneColumn':
      default:
        columns.push(this.getColumn(zoneIndex, sectionIndex++, 12));
        break;
    }

    return columns;
  }

  private getColumn(zoneIndex: number, sectionIndex: number, sectionFactor: number): Control {
    return {
      displayMode: 2,
      position: {
        zoneIndex: zoneIndex,
        sectionIndex: sectionIndex,
        sectionFactor: sectionFactor,
        layoutIndex: 1
      },
      emphasis: {}
    };
  }
}

module.exports = new SpoPageSectionAddCommand();