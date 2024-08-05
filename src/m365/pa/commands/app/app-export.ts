import fs from 'fs';
import path from 'path';
import { setTimeout } from 'timers/promises';
import GlobalOptions from '../../../../GlobalOptions.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  environmentName: string;
  packageDisplayName: string;
  packageDescription?: string;
  packageCreatedBy?: string;
  packageSourceEnvironment?: string;
  path?: string;
}

class PaAppExportCommand extends PowerPlatformCommand {
  private pollingInterval = 5000;

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
        option: '-n, --name <name>'
      },
      {
        option: '-e, --environmentName <environmentName>'
      },
      {
        option: '--packageDisplayName [packageDisplayName]'
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
        if (!validation.isValidGuid(args.options.name)) {
          return `${args.options.name} is not a valid GUID for option name`;
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

      path += `${filename}.zip`;

      fs.writeFileSync(path, file, 'binary');

      if (this.verbose) {
        await logger.logToStderr(`File saved to path '${path}'`);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getPackageResources(args: CommandArgs, logger: Logger): Promise<any> {
    if (this.verbose) {
      await logger.logToStderr('Getting the Microsoft Power App resources...');
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/providers/Microsoft.BusinessAppPlatform/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/listPackageResources?api-version=2016-11-01`,
      headers: {
        accept: 'application/json'
      },
      data: {
        baseResourceIds: [
          `/providers/Microsoft.PowerApps/apps/${args.options.name}`
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
      await logger.logToStderr(`Initiating package export for Microsoft Power App ${args.options.name}...`);
    }

    const resources = await this.getPackageResources(args, logger);

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/providers/Microsoft.BusinessAppPlatform/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/exportPackage?api-version=2016-11-01`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json',
      data: {
        includedResourceIds: [
          `/providers/Microsoft.PowerApps/apps/${args.options.name}`
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
      await logger.logToStderr('Retrieving the package link and waiting on the exported package.');
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
        await setTimeout(this.pollingInterval);
      }

      if (this.verbose) {
        await logger.logToStderr(`Current status of the get package link: ${status}`);
      }

    } while (status === 'Running');

    return link;
  }
}

export default new PaAppExportCommand();