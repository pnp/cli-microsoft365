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
  userId?: number;
  limit?: number;
}

class YammerGroupListCommand extends YammerCommand {
  private items: any[];

  constructor() {
    super();
    this.items = [];
  }

  public get name(): string {
    return `${commands.YAMMER_GROUP_LIST}`;
  }

  public get description(): string {
    return 'Returns the list of groups in a Yammer network or the groups for a specific user';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.userId = args.options.userId !== undefined;
    telemetryProps.limit = args.options.limit !== undefined;
    return telemetryProps;
  }

  private getAllItems(cmd: CommandInstance, args: CommandArgs, page: number): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      let endpoint = `${this.resource}/v1`;

      if (args.options.userId) {
        endpoint += `/groups/for_user/${args.options.userId}.json`
      }
      else {
        endpoint += `/groups.json`;
      }
      endpoint += `?page=${page}`;

      const requestOptions: any = {
        url: endpoint,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json;odata=nometadata'
        },
        json: true
      };

      request
        .get(requestOptions)
        .then((res: any): void => {
          this.items = this.items.concat(res);

          if (args.options.limit && this.items.length > args.options.limit) {
            this.items = this.items.slice(0, args.options.limit);
            resolve();
          }
          else {
            // we need to page by 50 items (hardcoded)
            if (this.items.length % 50 === 0) {
              this
                .getAllItems(cmd, args, ++page)
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
              name: n.name,
              email: n.email,
              privacy: n.privacy,
              external: n.external,
              moderated: n.moderated
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
        option: '--userId [userId]',
        description: 'Returns the groups for a specific user'
      },
      {
        option: '--limit [limit]',
        description: 'Limits the groups returned'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (args.options.userId && typeof args.options.userId !== 'number') {
        return `${args.options.userId} is not a number`;
      }

      if (args.options.limit && typeof args.options.limit !== 'number') {
        return `${args.options.limit} is not a number`;
      }

      return true;
    };
  }
}

module.exports = new YammerGroupListCommand();