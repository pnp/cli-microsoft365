import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import YammerCommand from '../../../base/YammerCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userId?: number;
  limit?: number;
}

class YammerGroupListCommand extends YammerCommand {
  private items: any[];

  public get name(): string {
    return commands.GROUP_LIST;
  }

  public get description(): string {
    return 'Returns the list of groups in a Yammer network or the groups for a specific user';
  }

  constructor() {
    super();
    this.items = [];

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        userId: args.options.userId !== undefined,
        limit: args.options.limit !== undefined
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--userId [userId]'
      },
      {
        option: '--limit [limit]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.userId && typeof args.options.userId !== 'number') {
          return `${args.options.userId} is not a number`;
        }

        if (args.options.limit && typeof args.options.limit !== 'number') {
          return `${args.options.limit} is not a number`;
        }

        return true;
      }
    );
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'name', 'email', 'privacy', 'external', 'moderated'];
  }

  private async getAllItems(logger: Logger, args: CommandArgs, page: number): Promise<void> {
    let endpoint = `${this.resource}/v1`;

    if (args.options.userId) {
      endpoint += `/groups/for_user/${args.options.userId}.json`;
    }
    else {
      endpoint += `/groups.json`;
    }
    endpoint += `?page=${page}`;

    const requestOptions: CliRequestOptions = {
      url: endpoint,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const output = await request.get<any[]>(requestOptions);
    if (!output.length) {
      return;
    }

    this.items = this.items.concat(output);

    if (args.options.limit && this.items.length > args.options.limit) {
      this.items = this.items.slice(0, args.options.limit);
    }
    else if (this.items.length % 50 === 0) {
      await this.getAllItems(logger, args, ++page);
    }
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    this.items = []; // this will reset the items array in interactive mode

    try {
      await this.getAllItems(logger, args, 1);
      logger.log(this.items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new YammerGroupListCommand();