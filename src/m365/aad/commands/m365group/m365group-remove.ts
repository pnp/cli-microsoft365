import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { aadGroup } from '../../../../utils/aadGroup.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import config from '../../../../config.js';
import { formatting } from '../../../../utils/formatting.js';
import { ClientSvcResponse, ClientSvcResponseContents, FormDigestInfo, spo } from '../../../../utils/spo.js';
import { setTimeout } from 'timers/promises';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  force?: boolean;
  skipRecycleBin: boolean;
}

class AadM365GroupRemoveCommand extends GraphCommand {
  private static maxRetries: number = 10;
  private intervalInMs: number = 5000;

  public get name(): string {
    return commands.M365GROUP_REMOVE;
  }

  public get description(): string {
    return 'Removes a Microsoft 365 Group';
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
        force: (!(!args.options.force)).toString(),
        skipRecycleBin: args.options.skipRecycleBin
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id <id>'
      },
      {
        option: '-f, --force'
      },
      {
        option: '--skipRecycleBin'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeGroup = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Removing Microsoft 365 Group: ${args.options.id}...`);
      }

      try {
        const isUnifiedGroup = await aadGroup.isUnifiedGroup(args.options.id);

        if (!isUnifiedGroup) {
          throw Error(`Specified group with id '${args.options.id}' is not a Microsoft 365 group.`);
        }

        const siteUrl = await this.getM365GroupSiteUrl(logger, args.options.id);
        const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);

        // Delete the Microsoft 365 group site. This operation will also delete the group.
        await this.deleteM365GroupSite(logger, siteUrl, spoAdminUrl);

        if (args.options.skipRecycleBin) {
          await this.deleteM365GroupFromRecycleBin(logger, args.options.id);
          await this.deleteSiteFromRecycleBin(logger, siteUrl, spoAdminUrl);
        }
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeGroup();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the group ${args.options.id}?`
      });

      if (result.continue) {
        await removeGroup();
      }
    }
  }

  private async getM365GroupSiteUrl(logger: Logger, id: string): Promise<string> {
    if (this.verbose) {
      await logger.logToStderr(`Getting the site URL of Microsoft 365 Group: ${id}...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/groups/${id}/drive?$select=webUrl`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const res = await request.get<{ webUrl: string }>(requestOptions);

    // Extract the base URL by removing everything after the last '/' character in the URL.
    const baseUrl = res.webUrl.substring(0, res.webUrl.lastIndexOf('/'));

    return baseUrl;
  }

  private async deleteM365GroupSite(logger: Logger, url: string, spoAdminUrl: string): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Deleting the group site: '${url}'...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${spoAdminUrl}/_api/GroupSiteManager/Delete?siteUrl='${url}'`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    await request.post(requestOptions);
  }

  private async deleteM365GroupFromRecycleBin(logger: Logger, id: string): Promise<void> {
    for (let retries = 0; retries < AadM365GroupRemoveCommand.maxRetries; retries++) {
      if (await this.isM365GroupInDeletedItemsList(id)) {
        await this.removeM365GroupPermanently(logger, id);
        return;
      }
      else {
        if (this.verbose) {
          await logger.logToStderr(`Group has not been deleted yet. Waiting for ${this.intervalInMs / 1000} seconds before the next attempt. ${AadM365GroupRemoveCommand.maxRetries - retries} attempts remaining...`);
        }

        await setTimeout(this.intervalInMs);
      }
    }

    await logger.logToStderr(`Group has been successfully deleted, but it couldn't be permanently removed from the recycle bin after all retries. It will still appear in the deleted groups list.`);
  }

  private async removeM365GroupPermanently(logger: Logger, id: string): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Group has been deleted and is now available in the deleted groups list. Removing permanently...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/directory/deletedItems/${id}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      }
    };

    await request.delete(requestOptions);
  }

  private async isM365GroupInDeletedItemsList(id: string): Promise<boolean> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/directory/deletedItems/${id}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      const response = await request.get<{ id: string }>(requestOptions);
      return Boolean(response && response.id);
    }
    catch (error: any) {
      if (error.response && error.response.status === 404) {
        return false;
      }
      else {
        throw error;
      }
    }
  }

  private async deleteSiteFromRecycleBin(logger: Logger, url: string, spoAdminUrl: string): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Deleting the M365 group site '${url}' from the recycle bin...`);
    }

    const res: FormDigestInfo = await spo.ensureFormDigest(spoAdminUrl as string, logger, undefined, this.debug);

    const requestOptions: CliRequestOptions = {
      url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': res.FormDigestValue
      },
      data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">${formatting.escapeXml(url)}</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
    };

    const processQuery: string = await request.post(requestOptions);
    const json: ClientSvcResponse = JSON.parse(processQuery);
    const response: ClientSvcResponseContents = json[0];

    if (response.ErrorInfo) {
      throw response.ErrorInfo.ErrorMessage;
    }
  }
}

export default new AadM365GroupRemoveCommand();