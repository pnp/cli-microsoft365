import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  groupId?: string;
  mailNickname?: string;
}

class AadO365GroupTeamifyCommand extends GraphCommand {
  public get name(): string {
    return commands.O365GROUP_TEAMIFY;
  }

  public get description(): string {
    return 'Creates a new Microsoft Teams team under existing Microsoft 365 group';
  }

  private getGroupId(args: CommandArgs): Promise<string> {
    if (args.options.groupId) {
      return Promise.resolve(args.options.groupId);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/groups?$filter=mailNickname eq '${encodeURIComponent(args.options.mailNickname as string)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: [{ id: string }] }>(requestOptions)
      .then(response => {
        const groupItem: { id: string } | undefined = response.value[0];

        if (!groupItem) {
          return Promise.reject(`The specified Microsoft 365 Group does not exist`);
        }

        if (response.value.length > 1) {
          return Promise.reject(`Multiple Microsoft 365 Groups with name ${args.options.mailNickname} found: ${response.value.map(x => x.id)}`);
        }

        return Promise.resolve(groupItem.id);
      });
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const data: any = {
      "memberSettings": {
        "allowCreatePrivateChannels": true,
        "allowCreateUpdateChannels": true
      },
      "messagingSettings": {
        "allowUserEditMessages": true,
        "allowUserDeleteMessages": true
      },
      "funSettings": {
        "allowGiphy": true,
        "giphyContentRating": "strict"
      }
    };

    this
      .getGroupId(args)
      .then((groupId: string): Promise<string> => {
        const requestOptions: any = {
          url: `${this.resource}/v1.0/groups/${encodeURIComponent(groupId)}/team`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          data: data,
          responseType: 'json'
        };

        return request.put(requestOptions);
      })
      .then(_ => cb(), (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --groupId [groupId]'
      },
      {
        option: '--mailNickname [mailNickname]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.groupId && args.options.mailNickname) {
      return 'Specify either groupId or mailNickname, but not both.';
    }

    if (!args.options.groupId && !args.options.mailNickname) {
      return 'Specify groupId or mailNickname, one is required';
    }

    if (args.options.groupId && !Utils.isValidGuid(args.options.groupId)) {
      return `${args.options.groupId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new AadO365GroupTeamifyCommand();