import request from '../../../../request';
import commands from '../../commands';
import { CommandOption, CommandValidate } from '../../../../Command';
import { v4 } from 'uuid';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import GlobalOptions from '../../../../GlobalOptions';
import {
  ClientSideWebpart,
  ClientSidePageComponent
} from './clientsidepages';
import { StandardWebPart, StandardWebPartUtils } from '../../StandardWebPartTypes';
import { isNumber } from 'util';
import { Control } from './canvasContent';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

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

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    let canvasContent: Control[];

    let pageFullName: string = args.options.pageName;
    if (args.options.pageName.indexOf('.aspx') < 0) {
      pageFullName += '.aspx';
    }

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
      .then((): Promise<ClientSideWebpart> => {
        if (this.verbose) {
          cmd.log(
            `Retrieving definition for web part ${args.options.webPartId ||
            args.options.standardWebPart}...`
          );
        }
        // Get the WebPart according to arguments
        return this.getWebPart(cmd, args);
      })
      .then((webPart: ClientSideWebpart): Promise<void> => {
        if (this.verbose) {
          cmd.log(`Setting client-side web part layout and properties...`);
        }

        this.setWebPartProperties(webPart, cmd, args);

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
      })
      .catch((err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  private getWebPart(cmd: CommandInstance, args: CommandArgs): Promise<any> {
    return new Promise<any>((resolve: (webPart: any) => void, reject: (error: any) => void): void => {
      const standardWebPart: string | undefined = args.options.standardWebPart;

      const webPartId = standardWebPart
        ? StandardWebPartUtils.getWebPartId(standardWebPart as StandardWebPart)
        : args.options.webPartId;

      if (this.debug) {
        cmd.log(`StandardWebPart: ${standardWebPart}`);
        cmd.log(`WebPartId: ${webPartId}`);
      }

      const requestOptions: any = {
        url: `${args.options.webUrl}/_api/web/getclientsidewebparts()`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        json: true
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
            cmd.log('WebPart definition:');
            cmd.log(webPartDefinition);
            cmd.log('');
          }

          if (this.verbose) {
            cmd.log(`Creating instance from definition of WebPart ${webPartId}...`);
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

  private setWebPartProperties(webPart: ClientSideWebpart, cmd: CommandInstance, args: CommandArgs): void {
    if (args.options.webPartProperties) {
      if (this.debug) {
        cmd.log('WebPart properties: ');
        cmd.log(args.options.webPartProperties);
        cmd.log('');
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
        cmd.log('WebPart data:');
        cmd.log(args.options.webPartData);
        cmd.log('');
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
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the page to add the web part to is located'
      },
      {
        option: '-n, --pageName <pageName>',
        description: 'Name of the page to which add the web part'
      },
      {
        option: '--standardWebPart [standardWebPart]',
        description: `Name of the standard web part to add (see the possible values below)`
      },
      {
        option: '--webPartId [webPartId]',
        description: 'ID of the custom web part to add'
      },
      {
        option: '--webPartProperties [webPartProperties]',
        description: 'JSON string with web part properties to set on the web part. Specify webPartProperties or webPartData but not both'
      },
      {
        option: '--webPartData [webPartData]',
        description: 'JSON string with web part data as retrieved from the web part maintenance mode. Specify webPartProperties or webPartData but not both'
      },
      {
        option: '--section [section]',
        description: 'Number of the section to which the web part should be added (1 or higher)'
      },
      {
        option: '--column [column]',
        description: 'Number of the column in which the web part should be added (1 or higher)'
      },
      {
        option: '--order [order]',
        description: 'Order of the web part in the column'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
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
    };
  }
}

module.exports = new SpoPageClientSideWebPartAddCommand();
