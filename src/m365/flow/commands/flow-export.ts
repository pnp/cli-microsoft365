import commands from '../commands';
import GlobalOptions from '../../../GlobalOptions';
import request from '../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../Command';
import AzmgmtCommand from '../../base/AzmgmtCommand';
import Utils from '../../../Utils';
import * as path from 'path';
import * as fs from 'fs';
import { CommandInstance } from '../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environment: string;
  id: string;
  packageDisplayName?: string;
  packageDescription?: string;
  packageCreatedBy?: string;
  packageSourceEnvironment?: string;
  format?: string;
  path?: string;
}

class FlowExportCommand extends AzmgmtCommand {
  public get name(): string {
    return commands.FLOW_EXPORT;
  }

  public get description(): string {
    return 'Exports the specified Microsoft Flow as a file';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.packageDisplayName = typeof args.options.packageDisplayName !== 'undefined';
    telemetryProps.packageDescription = typeof args.options.packageDescription !== 'undefined';
    telemetryProps.packageCreatedBy = typeof args.options.packageCreatedBy !== 'undefined';
    telemetryProps.packageSourceEnvironment = typeof args.options.packageSourceEnvironment !== 'undefined';
    telemetryProps.format = args.options.format;
    telemetryProps.path = typeof args.options.path !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    let filenameFromApi = '';
    const formatArgument = args.options.format ? args.options.format.toLowerCase() : '';

    if (this.verbose) {
      cmd.log(`Retrieving package resources for Microsoft Flow ${args.options.id}...`);
    }

    ((): Promise<any> => {
      if (formatArgument === 'json') {
        if (this.debug) {
          cmd.log('format = json, skipping listing package resources step');
        }

        return Promise.resolve();
      }

      const requestOptions: any = {
        url: `${this.resource}providers/Microsoft.BusinessAppPlatform/environments/${encodeURIComponent(args.options.environment)}/listPackageResources?api-version=2016-11-01`,
        headers: {
          accept: 'application/json'
        },
        body: {
          "baseResourceIds": [
            `/providers/Microsoft.Flow/flows/${args.options.id}`
          ]
        },
        json: true
      };

      return request.post(requestOptions);
    })()
      .then((res: any): Promise<{}> => {
        if (this.verbose) {
          cmd.log(`Initiating package export for Microsoft Flow ${args.options.id}...`);
        }

        const requestOptions: any = {
          url: `${this.resource}providers/${formatArgument === 'json' ?
            `Microsoft.ProcessSimple/environments/${encodeURIComponent(args.options.environment)}/flows/${encodeURIComponent(args.options.id)}?api-version=2016-11-01`
            : `Microsoft.BusinessAppPlatform/environments/${encodeURIComponent(args.options.environment)}/exportPackage?api-version=2016-11-01`}`,
          headers: {
            accept: 'application/json'
          },
          json: true
        };

        if (formatArgument !== 'json') {
          requestOptions['body'] = {
            "includedResourceIds": [
              `/providers/Microsoft.Flow/flows/${args.options.id}`
            ],
            "details": {
              "displayName": args.options.packageDisplayName,
              "description": args.options.packageDescription,
              "creator": args.options.packageCreatedBy,
              "sourceEnvironment": args.options.packageSourceEnvironment
            },
            "resources": res.resources
          }
        }

        return formatArgument === 'json' ? request.get(requestOptions) : request.post(requestOptions);
      })
      .then((res: any): Promise<string> => {
        if (this.verbose) {
          cmd.log(`Getting file for Microsoft Flow ${args.options.id}...`);
        }

        if (res.errors && res.errors.length && res.errors.length > 0) {
          return Promise.reject(res.errors[0].message)
        }

        const downloadFileUrl: string = formatArgument === 'json' ? '' : res.packageLink.value;
        const filenameRegEx: RegExp = /([^\/]+\.zip)/i;
        filenameFromApi = formatArgument === 'json' ? `${res.properties.displayName}.json` : (filenameRegEx.exec(downloadFileUrl) || ['output.zip'])[0];

        if (this.debug) {
          cmd.log(`Filename from PowerApps API: ${filenameFromApi}`);
          cmd.log('');
        }

        const requestOptions: any = {
          url: formatArgument === 'json' ?
            `${this.resource}/providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(args.options.environment)}/flows/${encodeURIComponent(args.options.id)}/exportToARMTemplate?api-version=2016-11-01`
            : downloadFileUrl,
          encoding: null, // Set encoding to null, otherwise binary data will be encoded to utf8 and binary data is corrupt 
          headers: formatArgument === 'json' ? {
            accept: 'application/json'
          } : {
            'x-anonymous': true
          }
        };

        return formatArgument === 'json' ?
          request.post(requestOptions)
          : request.get(requestOptions);
      })
      .then((file: string): void => {
        const path = args.options.path ? args.options.path : `./${filenameFromApi}`

        fs.writeFileSync(path, file, 'binary');
        if (!args.options.path || this.verbose) {
          if (this.verbose) {
            cmd.log(`File saved to path '${path}'`);
          }
          else {
            cmd.log(path);
          }
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>',
        description: 'The id of the Microsoft Flow to export'
      },
      {
        option: '-e, --environment <environment>',
        description: 'The name of the environment from which to export the Flow from'
      },
      {
        option: '-n, --packageDisplayName [packageDisplayName]',
        description: 'The display name to use in the exported package'
      },
      {
        option: '-d, --packageDescription [packageDescription]',
        description: 'The description to use in the exported package'
      },
      {
        option: '-c, --packageCreatedBy [packageCreatedBy]',
        description: 'The name of the person to be used as the creator of the exported package'
      },
      {
        option: '-s, --packageSourceEnvironment [packageSourceEnvironment]',
        description: 'The name of the source environment from which the exported package was taken'
      },
      {
        option: '-f, --format [format]',
        description: 'The format to export the Flow to json|zip. Default json'
      },
      {
        option: '-p, --path [path]',
        description: 'The path to save the exported package to'
      },

    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      const lowerCaseFormat = args.options.format ? args.options.format.toLowerCase() : '';

      if (!Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      if (args.options.format && (lowerCaseFormat !== 'json' && lowerCaseFormat !== 'zip')) {
        return 'Option format must be json or zip. Default is zip';
      }

      if (lowerCaseFormat === 'json') {
        if (args.options.packageCreatedBy) {
          return 'packageCreatedBy cannot be specified with output of json';
        }

        if (args.options.packageDescription) {
          return 'packageDescription cannot be specified with output of json';
        }

        if (args.options.packageDisplayName) {
          return 'packageDisplayName cannot be specified with output of json';
        }

        if (args.options.packageSourceEnvironment) {
          return 'packageSourceEnvironment cannot be specified with output of json';
        }
      }

      if (args.options.path && !fs.existsSync(path.dirname(args.options.path))) {
        return 'Specified path where to save the file does not exist';
      }

      return true;
    };
  }
}

module.exports = new FlowExportCommand();