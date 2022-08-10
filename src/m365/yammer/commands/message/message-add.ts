import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import YammerCommand from '../../../base/YammerCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  body: string;
  repliedToId?: number;
  directToUserIds?: string;
  groupId?: number;
  networkId?: number;
}

class YammerMessageAddCommand extends YammerCommand {
  public get name(): string {
    return commands.MESSAGE_ADD;
  }

  public get description(): string {
    return 'Posts a Yammer network message on behalf of the current user';
  }

  public defaultProperties(): string[] | undefined {
    return ['id'];
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
        repliedToId: args.options.repliedToId !== undefined,
        directToUserIds: args.options.directToUserIds !== undefined,
        groupId: args.options.groupId !== undefined,
        networkId: args.options.networkId !== undefined
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-b, --body <body>'
      },
      {
        option: '-r, --repliedToId [repliedToId]'
      },
      {
        option: '-d, --directToUserIds [directToUserIds]'
      },
      {
        option: '--groupId [groupId]'
      },
      {
        option: '--networkId [networkId]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.groupId && typeof args.options.groupId !== 'number') {
          return `${args.options.groupId} is not a number`;
        }

        if (args.options.networkId && typeof args.options.networkId !== 'number') {
          return `${args.options.networkId} is not a number`;
        }

        if (args.options.repliedToId && typeof args.options.repliedToId !== 'number') {
          return `${args.options.repliedToId} is not a number`;
        }

        if (args.options.groupId === undefined &&
          args.options.directToUserIds === undefined &&
          args.options.repliedToId === undefined) {
          return "You must either specify groupId, repliedToId or directToUserIds";
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const requestOptions: any = {
      url: `${this.resource}/v1/messages.json`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json;odata=nometadata'
      },
      responseType: 'json',
      data: {
        body: args.options.body,
        replied_to_id: args.options.repliedToId,
        direct_to_user_ids: args.options.directToUserIds,
        group_id: args.options.groupId,
        network_id: args.options.networkId
      }
    };

    request
      .post(requestOptions)
      .then((res: any): void => {
        let result = null;
        if (res.messages && res.messages.length === 1) {
          result = res.messages[0];
        }

        logger.log(result);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new YammerMessageAddCommand();
