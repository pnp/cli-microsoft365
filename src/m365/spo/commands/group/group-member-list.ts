import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  webUrl: string;
  groupId?: number;
  groupName?: string;
}

class SpoGroupMemberListCommand extends SpoCommand {
  public get name(): string {
    return commands.GROUP_MEMBER_LIST;
  }

  public get description(): string {
    return `List the members of a SharePoint Group`;
  }

  public defaultProperties(): string[] | undefined {
    return ['Title', 'UserPrincipalName', 'Id', 'Email'];
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
        groupId: typeof args.options.groupId !== 'undefined',
        groupName: typeof args.options.groupName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--groupId [groupId]'
      },
      {
        option: '--groupName [groupName]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.groupId && isNaN(args.options.groupId)) {
          return `Specified "groupId" ${args.options.groupId} is not valid`;
        }

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['groupName', 'groupId'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving the list of members from the SharePoint group :  ${args.options.groupId ? args.options.groupId : args.options.groupName}`);
    }

    const requestUrl: string = `${args.options.webUrl}/_api/web/sitegroups/${args.options.groupId
      ? `GetById('${args.options.groupId}')`
      : `GetByName('${formatting.encodeQueryParameter(args.options.groupName as string)}')`}/users`;

    try {
      const response = await odata.getAllItems<any>(requestUrl);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoGroupMemberListCommand();