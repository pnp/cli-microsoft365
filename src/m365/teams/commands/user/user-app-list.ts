import * as chalk from 'chalk';
import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import Utils from '../../../../Utils';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import request from '../../../../request';
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
    return `${commands.TEAMS_USER_APP_LIST}`;
  }

  public get description(): string {
    return 'Lists the apps deployed in the personal scope of the specified user';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    var userIdPromise: Promise<{ value: string; }>;

    if(args.options.userName) {
      const requestOptions: any = {
        url: `${this.resource}/v1.0/users/${encodeURIComponent(args.options.userName)}/id`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      userIdPromise = request.get<{ value: string; }>(requestOptions);  
    } else {
      userIdPromise = Promise.resolve({ value : args.options.userId });
    }

    userIdPromise.then((userId) => {
      const endpoint: string = `${this.resource}/v1.0/users/${encodeURIComponent(userId.value)}/teamwork/installedApps`
      
      this.getAllItems(endpoint, logger, true)
        .then((): void => {
          this.items.map(i => {
            var userAppId = Buffer.from(i.id, 'base64').toString('ascii');
            var appId = userAppId.substr(userAppId.indexOf("##") + 2, userAppId.length - userId.value.length - 2)
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
  
          if (this.verbose) {
            logger.log(chalk.green('DONE'));
          }
  
          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));   
    }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--userId',
        description: 'The ID of user to get the apps from'
      },
      {
        option: '--userName',
        description: 'The UPN of user to get the apps from'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if(!args.options.userId && !args.options.userName) {
      return `--userId or --userName need to be provided`;
    }

    if(args.options.userId && args.options.userName) {
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