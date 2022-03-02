import * as fs from 'fs';
import * as path from 'path';
import { Logger } from '../../../cli';
import {
  CommandOption
} from '../../../Command';
import GlobalOptions from '../../../GlobalOptions';
import request from '../../../request';
import { validation } from '../../../utils';
import AzmgmtCommand from '../../base/AzmgmtCommand';
import commands from '../commands';

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
    return commands.EXPORT;
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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    let filenameFromApi = '';
    const formatArgument = args.options.format ? args.options.format.toLowerCase() : '';

    if (this.verbose) {
      logger.logToStderr(`Retrieving package resources for Microsoft Flow ${args.options.id}...`);
    }

    ((): Promise<any> => {
      if (formatArgument === 'json') {
        if (this.debug) {
          logger.logToStderr('format = json, skipping listing package resources step');
        }

        return Promise.resolve();
      }

      const requestOptions: any = {
        url: `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/${encodeURIComponent(args.options.environment)}/listPackageResources?api-version=2016-11-01`,
        headers: {
          accept: 'application/json'
        },
        data: {
          "baseResourceIds": [
            `/providers/Microsoft.Flow/flows/${args.options.id}`
          ]
        },
        responseType: 'json'
      };

      return request.post(requestOptions);
    })()
      .then((res: any): Promise<any> => {
        if (typeof res !== 'undefined' && res.errors && res.errors.length && res.errors.length > 0) {
          return Promise.reject(res.errors[0].message);
        }

        if (this.verbose) {
          logger.logToStderr(`Initiating package export for Microsoft Flow ${args.options.id}...`);
        }

        const requestOptions: any = {
          url: `${this.resource}providers/${formatArgument === 'json' ?
            `Microsoft.ProcessSimple/environments/${encodeURIComponent(args.options.environment)}/flows/${encodeURIComponent(args.options.id)}?api-version=2016-11-01`
            : `Microsoft.BusinessAppPlatform/environments/${encodeURIComponent(args.options.environment)}/exportPackage?api-version=2016-11-01`}`,
          headers: {
            accept: 'application/json'
          },
          responseType: 'json'
        };

        if (formatArgument !== 'json') {
          // adds suggestedCreationType property to all resources
          // see https://github.com/pnp/cli-microsoft365/issues/1845
          Object.keys(res.resources).forEach((key) => {
            res.resources[key].type === 'Microsoft.Flow/flows'
              ? res.resources[key].suggestedCreationType = 'Update'
              : res.resources[key].suggestedCreationType = 'Existing';
          });

          requestOptions['data'] = {
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
          };
        }

        return formatArgument === 'json' ? request.get(requestOptions) : request.post(requestOptions);
      })
      .then((res: any): Promise<string> => {
        if (this.verbose) {
          logger.logToStderr(`Getting file for Microsoft Flow ${args.options.id}...`);
        }

        const downloadFileUrl: string = formatArgument === 'json' ? '' : res.packageLink.value;
        const filenameRegEx: RegExp = /([^\/]+\.zip)/i;
        filenameFromApi = formatArgument === 'json' ? `${res.properties.displayName}.json` : (filenameRegEx.exec(downloadFileUrl) || ['output.zip'])[0];

        if (this.debug) {
          logger.logToStderr(`Filename from PowerApps API: ${filenameFromApi}`);
          logger.logToStderr('');
        }

        const requestOptions: any = {
          url: formatArgument === 'json' ?
            `${this.resource}providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(args.options.environment)}/flows/${encodeURIComponent(args.options.id)}/exportToARMTemplate?api-version=2016-11-01`
            : downloadFileUrl,
          // Set responseType to arraybuffer, otherwise binary data will be encoded
          // to utf8 and binary data is corrupt
          responseType: 'arraybuffer',
          headers: formatArgument === 'json' ?
            {
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
        const path = args.options.path ? args.options.path : `./${filenameFromApi}`;

        fs.writeFileSync(path, file, 'binary');
        if (!args.options.path || this.verbose) {
          if (this.verbose) {
            logger.logToStderr(`File saved to path '${path}'`);
          }
          else {
            logger.log(path);
          }
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>'
      },
      {
        option: '-e, --environment <environment>'
      },
      {
        option: '-n, --packageDisplayName [packageDisplayName]'
      },
      {
        option: '-d, --packageDescription [packageDescription]'
      },
      {
        option: '-c, --packageCreatedBy [packageCreatedBy]'
      },
      {
        option: '-s, --packageSourceEnvironment [packageSourceEnvironment]'
      },
      {
        option: '-f, --format [format]'
      },
      {
        option: '-p, --path [path]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    const lowerCaseFormat = args.options.format ? args.options.format.toLowerCase() : '';

    if (!validation.isValidGuid(args.options.id)) {
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
  }
}

module.exports = new FlowExportCommand();
