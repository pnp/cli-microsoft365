import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { AssociatedGroupPropertiesCollection } from './AssociatedGroupPropertiesCollection.js';
import { GroupProperties } from './GroupProperties.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  associatedGroupsOnly: boolean;
}

class SpoGroupListCommand extends SpoCommand {
  public get name(): string {
    return commands.GROUP_LIST;
  }

  public get description(): string {
    return 'Lists all the groups within specific web';
  }

  public defaultProperties(): string[] | undefined {
    return ['Id', 'Title', 'LoginName', 'IsHiddenInUI', 'PrincipalType', 'Type'];
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
        associatedGroupsOnly: (!(!args.options.associatedGroupsOnly)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--associatedGroupsOnly'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.webUrl)
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving list of groups for specified web at ${args.options.webUrl}...`);
    }

    const baseUrl = `${args.options.webUrl}/_api/web`;

    try {
      if (!args.options.associatedGroupsOnly) {
        await this.getSiteGroups(baseUrl, logger);
      }
      else {
        await this.getAssociatedGroups(baseUrl, args.options, logger);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getSiteGroups(baseUrl: string, logger: Logger): Promise<void> {
    const groupProperties = await odata.getAllItems<GroupProperties>(`${baseUrl}/sitegroups`);
    await logger.log(groupProperties);
  }

  private async getAssociatedGroups(baseUrl: string, options: Options, logger: Logger): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: baseUrl + '?$expand=AssociatedOwnerGroup,AssociatedMemberGroup,AssociatedVisitorGroup&$select=AssociatedOwnerGroup,AssociatedMemberGroup,AssociatedVisitorGroup',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const groupProperties = await request.get<AssociatedGroupPropertiesCollection>(requestOptions);

    if (!options.output || !Cli.shouldTrimOutput(options.output)) {
      await logger.log(groupProperties);
    }
    else {
      //converted to text friendly output
      const output = Object.getOwnPropertyNames(groupProperties).map(prop => ({ Type: prop, ...(groupProperties as any)[prop] }));
      await logger.log(output);
    }
  }
}

export default new SpoGroupListCommand();