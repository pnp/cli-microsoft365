import { CommandOption, CommandValidate } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import YammerCommand from '../../../base/YammerCommand';
import commands from '../../commands';
import { CommandInstance } from '../../../../cli';

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

  constructor() {
    super();
    this.items = [];
  }

  public get name(): string {
    return `${commands.YAMMER_USER_LIST}`;
  }

  public get description(): string {
    return 'Returns users from the current network';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.letter = args.options.letter !== undefined;
    telemetryProps.sortBy = args.options.sortBy !== undefined;
    telemetryProps.reverse = args.options.reverse !== undefined;
    telemetryProps.limit = args.options.limit !== undefined;
    telemetryProps.groupId = args.options.groupId !== undefined;
    return telemetryProps;
  }

  private getAllItems(cmd: CommandInstance, args: CommandArgs, page: number): Promise<void> {
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
        json: true
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
              this.getAllItems(cmd, args, ++page)
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
  };

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    this.items = []; // this will reset the items array in interactive mode

    this
      .getAllItems(cmd, args, 1)
      .then((): void => {
        if (args.options.output === 'json') {
          cmd.log(this.items);
        }
        else {
          cmd.log(this.items.map((n: any) => {
            const item: any = {
              id: n.id,
              full_name: n.full_name,
              email: n.email
            };
            return item;
          }));
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  };

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-g, --groupId [groupId]',
        description: 'Returns users within a given group'
      },
      {
        option: '-l, --letter [letter]',
        description: 'Returns users with usernames beginning with the given character'
      },
      {
        option: '--reverse',
        description: 'Returns users in reverse sorting order'
      },
      {
        option: '--limit [limit]',
        description: 'Limits the users returned'
      },
      {
        option: '--sortBy [sortBy]',
        description: 'Returns users sorted by a number of messages or followers, instead of the default behavior of sorting alphabetically. Allowed values are messages,followers',
        autocomplete: ['messages', 'followers']
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {

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
    };
  }
}

module.exports = new YammerUserListCommand();
