import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils/validation';
import { aadGroup } from '../../../../utils/aadGroup';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { GroupExtended } from './GroupExtended';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  includeSiteUrl: boolean;
}

class AadO365GroupGetCommand extends GraphCommand {
  public get name(): string {
    return commands.O365GROUP_GET;
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
      group = await aadGroup.getGroupById(args.options.id);

      if (args.options.includeSiteUrl) {
        const requestOptions: any = {
          url: `${this.resource}/v1.0/groups/${group.id}/drive?$select=webUrl`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        const res = await request.get<{ webUrl: string }>(requestOptions);
        group.siteUrl = res.webUrl ? res.webUrl.substr(0, res.webUrl.lastIndexOf('/')) : '';
      }

      logger.log(group);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new AadO365GroupGetCommand();
