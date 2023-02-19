import { Group } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils/validation';
import { aadGroup } from '../../../../utils/aadGroup';
import GraphCommand from '../../../base/GraphCommand';
import { Channel } from '../../Channel';
import commands from '../../commands';
import { formatting } from '../../../../utils/formatting';

interface ExtendedGroup extends Group {
  resourceProvisioningOptions: string[];
}

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId?: string;
  teamName?: string;
  id?: string;
  name?: string;
  primary?: boolean;
}

class TeamsChannelGetCommand extends GraphCommand {
  private teamId: string = "";

  public get name(): string {
    return commands.CHANNEL_GET;
  }

  public get description(): string {
    return 'Gets information about the specific Microsoft Teams team channel';
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
        teamId: typeof args.options.teamId !== 'undefined',
        teamName: typeof args.options.teamName !== 'undefined',
        id: typeof args.options.id !== 'undefined',
        name: typeof args.options.name !== 'undefined',
        primary: (!(!args.options.primary)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--teamId [teamId]'
      },
      {
        option: '--teamName [teamName]'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '--name [name]'
      },
      {
        option: '--primary'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
          return `${args.options.teamId} is not a valid GUID`;
        }

        if (args.options.id && !validation.isValidTeamsChannelId(args.options.id)) {
          return `${args.options.id} is not a valid Teams channel id`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['teamId', 'teamName'] },
      { options: ['id', 'name', 'primary'] }
    );
  }

  private getTeamId(args: CommandArgs): Promise<string> {
    if (args.options.teamId) {
      return Promise.resolve(args.options.teamId);
    }

    return aadGroup
      .getGroupByDisplayName(args.options.teamName!)
      .then(group => {
        if ((group as ExtendedGroup).resourceProvisioningOptions.indexOf('Team') === -1) {
          return Promise.reject(`The specified team does not exist in the Microsoft Teams`);
        }

        return group.id!;
      });
  }

  private getChannelId(args: CommandArgs): Promise<string> {
    if (args.options.id) {
      return Promise.resolve(args.options.id);
    }

    if (args.options.primary) {
      return Promise.resolve('');
    }

    const channelRequestOptions: any = {
      url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(this.teamId)}/channels?$filter=displayName eq '${formatting.encodeQueryParameter(args.options.name as string)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: Channel[] }>(channelRequestOptions)
      .then(response => {
        const channelItem: Channel | undefined = response.value[0];

        if (!channelItem) {
          return Promise.reject(`The specified channel does not exist in the Microsoft Teams team`);
        }

        return Promise.resolve(channelItem.id);
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      this.teamId = await this.getTeamId(args);
      const channelId: string = await this.getChannelId(args);
      let url: string = '';
      if (args.options.primary) {
        url = `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(this.teamId)}/primaryChannel`;
      }
      else {
        url = `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(this.teamId)}/channels/${formatting.encodeQueryParameter(channelId)}`;
      }

      const requestOptions: any = {
        url: url,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const res: Channel = await request.get<Channel>(requestOptions);
      logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new TeamsChannelGetCommand();