import request from '../../../../request';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import GlobalOptions from '../../../../GlobalOptions';
import { CanvasSectionTemplate } from './clientsidepages';
import { isNumber } from 'util';
import { Control } from './canvasContent';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  webUrl: string;
  sectionTemplate: string;
  order?: number;
}

class SpoPageSectionAddCommand extends SpoCommand {
  public get name(): string {
    return `${commands.PAGE_SECTION_ADD}`;
  }

  public get description(): string {
    return 'Adds section to modern page';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    let pageFullName: string = args.options.name.toLowerCase();
    if (pageFullName.indexOf('.aspx') < 0) {
      pageFullName += '.aspx';
    }
    let canvasContent: Control[];

    if (this.verbose) {
      cmd.log(`Retrieving page information...`);
    }

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/sitepages/pages/GetByUrl('sitepages/${encodeURIComponent(pageFullName)}')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      json: true
    };

    request
      .get<{ CanvasContent1: string; IsPageCheckedOutToCurrentUser: boolean }>(requestOptions)
      .then((res: { CanvasContent1: string; IsPageCheckedOutToCurrentUser: boolean }): Promise<void> => {
        canvasContent = JSON.parse(res.CanvasContent1 || "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]");

        if (res.IsPageCheckedOutToCurrentUser) {
          return Promise.resolve();
        }

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/sitepages/pages/GetByUrl('sitepages/${encodeURIComponent(pageFullName)}')/checkoutpage`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          json: true
        };

        return request.post(requestOptions);
      })
      .then((): Promise<void> => {
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

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/sitepages/pages/GetByUrl('sitepages/${encodeURIComponent(pageFullName)}')/savepage`,
          headers: {
            'accept': 'application/json;odata=nometadata',
            'content-type': 'application/json;odata=nometadata'
          },
          body: {
            CanvasContent1: JSON.stringify(canvasContent)
          },
          json: true
        };

        return request.post(requestOptions);
      })
      .then((): void => {
        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();

      }, (err: any): void => {
        this.handleRejectedODataJsonPromise(err, cmd, cb)
      });
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

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --name <name>',
        description: 'Name of the page to add section to'
      },
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the page to retrieve is located'
      },
      {
        option: '-t, --sectionTemplate <sectionTemplate>',
        description: 'Type of section to add. Allowed values OneColumn|OneColumnFullWidth|TwoColumn|ThreeColumn|TwoColumnLeft|TwoColumnRight'
      },
      {
        option: '--order [order]',
        description: 'Order of the section to add'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!(args.options.sectionTemplate in CanvasSectionTemplate)) {
        return `${args.options.sectionTemplate} is not a valid section template. Allowed values are OneColumn|OneColumnFullWidth|TwoColumn|ThreeColumn|TwoColumnLeft|TwoColumnRight`;
      }

      if (typeof args.options.order !== 'undefined') {
        if (!isNumber(args.options.order) || args.options.order < 1) {
          return 'The value of parameter order must be 1 or higher';
        }
      }

      return SpoCommand.isValidSharePointUrl(args.options.webUrl);
    };
  }
}

module.exports = new SpoPageSectionAddCommand();