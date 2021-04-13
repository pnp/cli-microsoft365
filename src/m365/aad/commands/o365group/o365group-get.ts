import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import teamsCommands from '../../../teams/commands';
import { Group } from './Group';

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

  public alias(): string[] | undefined {
    return [teamsCommands.TEAM_GET];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    let group: Group;

    const requestOptions: any = {
      url: `${this.resource}/v1.0/groups/${args.options.id}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    request
      .get<Group>(requestOptions)
      .then((res: Group): Promise<{ webUrl: string }> => {
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

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>'
      },
      {
        option: '--includeSiteUrl'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!Utils.isValidGuid(args.options.id)) {
      return `${args.options.id} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new AadO365GroupGetCommand();
