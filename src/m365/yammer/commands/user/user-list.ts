import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import YammerCommand from '../../../base/YammerCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  letter?: string;
  sortBy?: string;
  reverse?: boolean;
  limit?: number;
  groupId?: number;
}

class YammerUserListCommand extends YammerCommand {
  protected items: any[];

  public get name(): string {
    return commands.USER_LIST;
  }

  public get description(): string {
    return 'Returns users from the current network';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'full_name', 'email'];
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
        letter: args.options.letter !== undefined,
        sortBy: args.options.sortBy !== undefined,
        reverse: args.options.reverse !== undefined,
        limit: args.options.limit !== undefined,
        groupId: args.options.groupId !== undefined
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-g, --groupId [groupId]'
      },
      {
        option: '-l, --letter [letter]'
      },
      {
        option: '--reverse'
      },
      {
        option: '--limit [limit]'
      },
      {
        option: '--sortBy [sortBy]',
        autocomplete: ['messages', 'followers']
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.groupId && typeof args.options.groupId !== 'number') {
          return `${args.options.groupId} is not a number`;
        }

        if (args.options.limit && typeof args.options.limit !== 'number') {
          return `${args.options.limit} is not a number`;
        }

        if (args.options.sortBy && args.options.sortBy !== 'messages' && args.options.sortBy !== 'followers') {
          return `sortBy accepts only the values "messages" or "followers"`;
        }

        if (args.options.letter && !/^(?!\d)[a-zA-Z]+$/i.test(args.options.letter)) {
          return `Value of 'letter' is invalid. Only characters within the ranges [A - Z], [a - z] are allowed.`;
        }

        if (args.options.letter && args.options.letter.length !== 1) {
          return `Only one char as value of 'letter' accepted.`;
        }

        return true;
      }
    );
  }

  private getAllItems(logger: Logger, args: CommandArgs, page: number): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      if (page === 1) {
        this.items = [];
      }

      let endPoint = `${this.resource}/v1/users.json`;

      if (args.options.groupId !== undefined) {
        endPoint = `${this.resource}/v1/users/in_group/${args.options.groupId}.json`;
      }

      endPoint += `?page=${page}`;
      if (args.options.reverse !== undefined) {
        endPoint += `&reverse=true`;
      }
      if (args.options.sortBy !== undefined) {
        endPoint += `&sort_by=${args.options.sortBy}`;
      }
      if (args.options.letter !== undefined) {
        endPoint += `&letter=${args.options.letter}`;
      }

      const requestOptions: any = {
        url: endPoint,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      request
        .get(requestOptions)
        .then((res: any): void => {
          let userOutput = res;
          // groups user retrieval returns a user array containing the user objects
          if (res.users) {
            userOutput = res.users;
          }

          this.items = this.items.concat(userOutput);

          // this is executed once at the end if the limit operation has been executed
          // we need to return the array of the desired size. The API does not provide such a feature
          if (args.options.limit !== undefined && this.items.length > args.options.limit) {
            this.items = this.items.slice(0, args.options.limit);
            resolve();
          }
          else {
            // if the groups endpoint is used, the more_available will tell if a new retrieval is required
            // if the user endpoint is used, we need to page by 50 items (hardcoded)
            if (res.more_available === true || this.items.length % 50 === 0) {
              this.getAllItems(logger, args, ++page)
                .then((): void => {
                  resolve();
                }, (err: any): void => {
                  reject(err);
                });
            }
            else {
              resolve();
            }
          }
        }, (err: any): void => {
          reject(err);
        });
    });
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this.items = []; // this will reset the items array in interactive mode

    this
      .getAllItems(logger, args, 1)
      .then((): void => {
        logger.log(this.items);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new YammerUserListCommand();
