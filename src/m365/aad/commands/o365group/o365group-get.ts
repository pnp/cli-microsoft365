import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    let group: GroupExtended;

    aadGroup
      .getGroupById(args.options.id)
      .then((res: GroupExtended): Promise<{ webUrl: string }> => {
        group = res;

        if (args.options.includeSiteUrl) {
          const requestOptions: any = {
            url: `${this.resource}/v1.0/groups/${group.id}/drive?$select=webUrl`,
            headers: {
              accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
          };

          return request.get(requestOptions);
        }
        else {
          return Promise.resolve(undefined as any);
        }
      })
      .then((res?: { webUrl: string }): void => {
        if (res) {
          group.siteUrl = res.webUrl ? res.webUrl.substr(0, res.webUrl.lastIndexOf('/')) : '';
        }

        logger.log(group);
        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new AadO365GroupGetCommand();
