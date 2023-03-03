import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { validation } from '../../../../utils/validation';
import commands from '../../commands';
import * as fs from 'fs';
import * as path from 'path';
import { formatting } from '../../../../utils/formatting';
import request, { CliRequestOptions } from '../../../../request';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  environment: string;
  packageDisplayName: string;
  packageDescription?: string;
  packageCreatedBy?: string;
  packageSourceEnvironment?: string;
  path?: string;
}

class PaAppExportCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.APP_EXPORT;
  }

  public get description(): string {
    return 'Exports the specified Power App';
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
        packageDescription: typeof args.options.packageDescription !== 'undefined',
        packageCreatedBy: typeof args.options.packageCreatedBy !== 'undefined',
        packageSourceEnvironment: typeof args.options.packageSourceEnvironment !== 'undefined',
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
        option: '-p, --path [path]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        if (args.options.path && !fs.existsSync(path.dirname(args.options.path))) {
          return 'Specified path where to save the file does not exist';
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const location = await this.exportPackage(args, logger);
      const packageLink = await this.getPackageLink(args, logger, location);
      //Replace all illegal characters from the file name
      const illegalCharsRegEx = /[\\\/:*?"<>|]/g;
      const filename = args.options.packageDisplayName.replace(illegalCharsRegEx, '_');

      const requestOptions: CliRequestOptions = {
        url: packageLink,
        // Set responseType to arraybuffer, otherwise binary data will be encoded
        // to utf8 and binary data is corrupt
        responseType: 'arraybuffer',
        headers: {
          'x-anonymous': true
        }
      };

      const file = await request.get<string>(requestOptions);

      let path = args.options.path || './';
      if (!path.endsWith('/')) {
        path += '/';
      }

      path += `${args.options.packageDisplayName}.zip`;

      fs.writeFileSync(path, file, 'binary');
      if (this.verbose) {
        logger.logToStderr(`File '${filename}' saved to path '${path}'`);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getPackageResources(args: CommandArgs, logger: Logger): Promise<any> {
    if (this.verbose) {
      logger.logToStderr('Getting the Microsoft Power App resources...');
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/providers/Microsoft.BusinessAppPlatform/environments/${formatting.encodeQueryParameter(args.options.environment)}/listPackageResources?api-version=2016-11-01`,
      headers: {
        accept: 'application/json'
      },
      data: {
        baseResourceIds: [
          `/providers/Microsoft.PowerApps/apps/${args.options.id}`
        ]
      },
      responseType: 'json'
    };

    const response = await request.post<any>(requestOptions);
    Object.keys(response.resources).forEach((key) => {
      response.resources[key].suggestedCreationType = 'Update';
    });
    return response.resources;
  }

  private async exportPackage(args: CommandArgs, logger: Logger): Promise<string> {
    if (this.verbose) {
      logger.logToStderr(`Initiating package export for Microsoft Power App ${args.options.id}...`);
    }

    const resources = await this.getPackageResources(args, logger);

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/providers/Microsoft.BusinessAppPlatform/environments/${formatting.encodeQueryParameter(args.options.environment)}/exportPackage?api-version=2016-11-01`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json',
      data: {
        includedResourceIds: [
          `/providers/Microsoft.PowerApps/apps/${args.options.id}`
        ],
        details: {
          creator: args.options.packageCreatedBy,
          description: args.options.packageDescription,
          displayName: args.options.packageDisplayName,
          sourceEnvironment: args.options.packageSourceEnvironment
        },
        resources: resources
      },
      fullResponse: true
    };

    const response: any = await request.post<any>(requestOptions);

    return response.headers.location;
  }

  private async getPackageLink(args: CommandArgs, logger: Logger, location: string): Promise<string> {
    if (this.verbose) {
      logger.logToStderr('Retrieving the package link and waiting on the exported package.');
    }

    let status;
    let link;

    const requestOptions: CliRequestOptions = {
      url: location,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    do {
      const response = await request.get<any>(requestOptions);
      status = response.properties.status;
      if (status === "Succeeded") {
        link = response.properties.packageLink.value;
      }
      else {
        await this.sleep(5000);
      }

      if (this.verbose) {
        logger.logToStderr(`Current status of the get package link: ${status}`);
      }

    } while (status === 'Running');

    return link;
  }

  protected sleep(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
}

module.exports = new PaAppExportCommand();