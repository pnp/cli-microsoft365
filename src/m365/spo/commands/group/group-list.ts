import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { odata } from '../../../../utils/odata';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { AssociatedGroupPropertiesCollection } from './AssociatedGroupPropertiesCollection';
import { GroupProperties } from './GroupProperties';

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
      logger.logToStderr(`Retrieving list of groups for specified web at ${args.options.webUrl}...`);
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
    logger.log(groupProperties);
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
    logger.log(groupProperties);
    if (!options.output || options.output === 'json') {
      logger.log(groupProperties);
    }
    else {
      //converted to text friendly output
      const output = Object.getOwnPropertyNames(groupProperties).map(prop => ({ Type: prop, ...(groupProperties as any)[prop] }));
      logger.log(output);
    }
  }
}

module.exports = new SpoGroupListCommand();