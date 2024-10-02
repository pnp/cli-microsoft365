import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { validation } from '../../../../utils/validation.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
import { PpWebSiteOptions } from '../Website.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environmentName: string;
  id?: string;
  websiteId?: string;
  websiteName?: string;
  asAdmin?: boolean;
  force?: boolean;
}

class PpWebSiteWebRoleRemoveCommand extends PowerPlatformCommand {

  public get name(): string {
    return commands.WEBSITE_WEBROLE_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified webrole';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: typeof args.options.id !== 'undefined',
        websiteId: typeof args.options.websiteId !== 'undefined',
        websiteName: typeof args.options.name !== 'undefined',
        asAdmin: !!args.options.asAdmin,
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-e, --environmentName <environmentName>'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '--websiteId [websiteId]'
      },
      {
        option: '--websiteName [websiteName]'
      },
      {
        option: '--asAdmin'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['websiteId', 'websiteName'] }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }
        if (args.options.websiteId && !validation.isValidGuid(args.options.websiteId)) {
          return `${args.options.websiteId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: any): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Removing websrole '${args.options.id}'...`);
    }

    if (args.options.force) {
      await this.deleteWebSiteWebRole(args);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove webrole '${args.options.id}'?` });

      if (result) {
        await this.deleteWebSiteWebRole(args);
      }
    }
  }

  private async getWebSiteId(dynamicsApiUrl: string, args: CommandArgs): Promise<any> {
    if (args.options.websiteId) {
      return args.options.websiteId;
    }

    const options: PpWebSiteOptions = {
      environmentName: args.options.environmentName,
      id: args.options.websiteId,
      name: args.options.websiteName,
      output: 'json'
    };

    const requestOptions: CliRequestOptions = {
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    if (options.name) {
      requestOptions.url = `${dynamicsApiUrl}/api/data/v9.2/powerpagesites?$filter=name eq '${options.name}'&$select=powerpagesiteid`;
      const result = await request.get<{ value: any[] }>(requestOptions);

      if (result.value.length === 0) {
        throw `The specified website '${args.options.websiteName}' does not exist.`;
      }
      return result.value[0].powerpagesiteid;
    }

  }

  private async getWebSiteWebRoleId(dynamicsApiUrl: string, args: CommandArgs): Promise<any> {

    const requestOptions: CliRequestOptions = {
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const websiteId = await this.getWebSiteId(dynamicsApiUrl, args);

    if (websiteId) {
      requestOptions.url = `${dynamicsApiUrl}/api/data/v9.2/mspp_webroles?$filter=_mspp_websiteid_value eq ${websiteId} and mspp_webroleid eq ${args.options.id}&$select=mspp_webroleid`;

      const result = await request.get<{ value: any[] }>(requestOptions);

      if (result.value.length === 0) {
        throw `The specified webrole '${args.options.id}' does not exist for the specified website.`;
      }
      return result.value[0].mspp_webroleid;

    }
  }

  private async deleteWebSiteWebRole(args: CommandArgs): Promise<void> {
    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);

      const websitewebroleId = await this.getWebSiteWebRoleId(dynamicsApiUrl, args);
      const requestOptions: CliRequestOptions = {
        url: `${dynamicsApiUrl}/api/data/v9.2/mspp_webroles(${websitewebroleId})`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      await request.delete(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PpWebSiteWebRoleRemoveCommand();