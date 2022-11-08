import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils/odata';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  groupId?: number;
  groupName?: string;
}

class SpoGroupUserListCommand extends SpoCommand {
  public get name(): string {
    return commands.GROUP_USER_LIST;
  }

  public get description(): string {
    return `List members of a SharePoint Group`;
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
    this.optionSets.push(['groupName', 'groupId']);
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving the list of members from the SharePoint group :  ${args.options.groupId ? args.options.groupId : args.options.groupName}`);
    }

    const requestUrl: string = `${args.options.webUrl}/_api/web/sitegroups/${args.options.groupId
      ? `GetById('${encodeURIComponent(args.options.groupId)}')`
      : `GetByName('${encodeURIComponent(args.options.groupName as string)}')`}/users`;

    try {
      const response = await odata.getAllItems<any>(requestUrl);
      logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoGroupUserListCommand();