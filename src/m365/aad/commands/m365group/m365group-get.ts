import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { aadGroup } from '../../../../utils/aadGroup.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { GroupExtended } from './GroupExtended.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  includeSiteUrl: boolean;
}

class AadM365GroupGetCommand extends GraphCommand {
  public get name(): string {
    return commands.M365GROUP_GET;
  }

  public get description(): string {
    return 'Gets information about the specified Microsoft 365 Group or Microsoft Teams team';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id <id>'
      },
      {
        option: '--includeSiteUrl'
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
    let group: GroupExtended;

    try {
      const isUnifiedGroup = await aadGroup.isUnifiedGroup(args.options.id);
      if (!isUnifiedGroup) {
        throw Error(`Specified group with id '${args.options.id}' is not a Microsoft 365 group.`);
      }

      group = await aadGroup.getGroupById(args.options.id);

      if (args.options.includeSiteUrl) {
        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/groups/${group.id}/drive?$select=webUrl`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        const res = await request.get<{ webUrl: string }>(requestOptions);
        group.siteUrl = res.webUrl ? res.webUrl.substr(0, res.webUrl.lastIndexOf('/')) : '';
      }

      await logger.log(group);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new AadM365GroupGetCommand();
