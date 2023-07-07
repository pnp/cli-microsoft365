import * as fs from 'fs';
import * as path from 'path';
import { Logger } from '../../../cli/Logger';
import GlobalOptions from '../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../request';
import { formatting } from '../../../utils/formatting';
import { validation } from '../../../utils/validation';
import PowerPlatformCommand from '../../base/PowerPlatformCommand';
import commands from '../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environmentName: string;
  id: string;
  packageDisplayName?: string;
  packageDescription?: string;
  packageCreatedBy?: string;
  packageSourceEnvironment?: string;
  format?: string;
  path?: string;
}

class FlowExportCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.EXPORT;
  }

  public get description(): string {
    return 'Exports the specified Microsoft Flow as a file';
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
        packageDisplayName: typeof args.options.packageDisplayName !== 'undefined',
        packageDescription: typeof args.options.packageDescription !== 'undefined',
        packageCreatedBy: typeof args.options.packageCreatedBy !== 'undefined',
        packageSourceEnvironment: typeof args.options.packageSourceEnvironment !== 'undefined',
        format: args.options.format,
        path: typeof args.options.path !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id <id>'
      },
      {
        option: '-e, --environmentName <environmentName>'
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
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
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
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let filenameFromApi = '';
    const formatArgument = args.options.format ? args.options.format.toLowerCase() : '';

    if (this.verbose) {
      logger.logToStderr(`Retrieving package resources for Microsoft Flow ${args.options.id}...`);
    }

    try {
      let res: any;
      if (formatArgument === 'json') {
        if (this.debug) {
          logger.logToStderr('format = json, skipping listing package resources step');
        }
      }
      else {
        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/providers/Microsoft.BusinessAppPlatform/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/listPackageResources?api-version=2016-11-01`,
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

        res = await request.post<any>(requestOptions);
      }

      if (typeof res !== 'undefined' && res.errors && res.errors.length && res.errors.length > 0) {
        throw res.errors[0].message;
      }

      if (this.verbose) {
        logger.logToStderr(`Initiating package export for Microsoft Flow ${args.options.id}...`);
      }

      let requestOptions: CliRequestOptions = {
        url: formatArgument === 'json' ?
          `https://management.azure.com/providers/Microsoft.ProcessSimple/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.id)}?api-version=2016-11-01`
          : `${this.resource}/providers/Microsoft.BusinessAppPlatform/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/exportPackage?api-version=2016-11-01`,
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

      res = formatArgument === 'json' ? await request.get(requestOptions) : await request.post(requestOptions);

      if (this.verbose) {
        logger.logToStderr(`Getting file for Microsoft Flow ${args.options.id}...`);
      }

      const downloadFileUrl: string = formatArgument === 'json' ? '' : res.packageLink.value;
      const filenameRegEx: RegExp = /([^\/]+\.zip)/i;
      filenameFromApi = formatArgument === 'json' ? `${res.properties.displayName}.json` : (filenameRegEx.exec(downloadFileUrl) || ['output.zip'])[0];
      // Replace all illegal characters from the file name
      const illegalCharsRegEx = /[\\\/:*?"<>|]/g;
      filenameFromApi = filenameFromApi.replace(illegalCharsRegEx, '_');

      if (this.debug) {
        logger.logToStderr(`Filename from PowerApps API: ${filenameFromApi}`);
        logger.logToStderr('');
      }

      requestOptions = {
        url: formatArgument === 'json' ?
          `https://management.azure.com/providers/Microsoft.ProcessSimple/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.id)}/exportToARMTemplate?api-version=2016-11-01`
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

      const file = formatArgument === 'json' ?
        await request.post<string>(requestOptions)
        : await request.get<string>(requestOptions);

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
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new FlowExportCommand();