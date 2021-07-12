import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import commands from '../../commands';
import { UserTeamsApp } from '../../UserTeamsApp';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userId: string;
  userName: string;
}

class TeamsUserAppListCommand extends GraphItemsListCommand<UserTeamsApp> {
  public get name(): string {
    return commands.USER_APP_LIST;
  }

  public get description(): string {
    return 'List the apps installed in the personal scope of the specified user';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.userId = typeof args.options.userId !== 'undefined';
    telemetryProps.userName = typeof args.options.userName !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    let userId: string = '';

    this
      .getUserId(args)
      .then((_userId): Promise<void> => {
        userId = _userId.value;
        const endpoint: string = `${this.resource}/v1.0/users/${encodeURIComponent(userId)}/teamwork/installedApps`;

        return this.getAllItems(endpoint, logger, true);
      })
      .then((): void => {
        this.items.map(i => {
          const userAppId: string = Buffer.from(i.id, 'base64').toString('ascii');
          const appId: string = userAppId.substr(userAppId.indexOf("##") + 2, userAppId.length - userId.length - 2);
          i.appId = appId;
        });
        
        if (args.options.output === 'json') {
          logger.log(this.items);
        }
        else {
          logger.log(this.items.map(i => {
            return {
              id: i.id,
              appId: i.appId
            };
          }));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getUserId(args: CommandArgs): Promise<{ value: string }> {
    if (args.options.userId) {
      return Promise.resolve({ value: args.options.userId });
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/users/${encodeURIComponent(args.options.userName)}/id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.get<{ value: string; }>(requestOptions);
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--userId'
      },
      {
        option: '--userName'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.userId && !args.options.userName) {
      return `--userId or --userName need to be provided`;
    }

    if (args.options.userId && args.options.userName) {
      return `Please specify either --userId or --userName, not both`;
    }

    if (args.options.userId && !Utils.isValidGuid(args.options.userId)) {
      return `${args.options.userId} is not a valid GUID`;
    }

    if (args.options.userName && !Utils.isValidUserPrincipalName(args.options.userName)) {
      return `${args.options.userName} is not a valid userName`;
    }

    return true;
  }
}

module.exports = new TeamsUserAppListCommand();