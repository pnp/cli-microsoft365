import { Cli, CommandOutput, Logger } from '../../../../cli';
import Command, { CommandErrorWithOutput, CommandError, CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import * as AadUserGetCommand from '../../../aad/commands/user/user-get';
import { Options as AadUserGetCommandOptions } from '../../../aad/commands/user/user-get';
import commands from '../../commands';
import { SharingResult } from './SharingResult';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  groupId?: number;
  groupName?: string;
  userName?: string;
  email?: string;
}

class SpoGroupUserAddCommand extends SpoCommand {
  private groupId: number = 0;

  public get name(): string {
    return commands.GROUP_USER_ADD;
  }

  public get description(): string {
    return 'Add a user or multiple users to SharePoint Group';
  }

  public defaultProperties(): string[] | undefined {
    return ['DisplayName', 'Email'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    this
      .getGroupId(args)
      .then((groupId: number): Promise<string[]> => {
        this.groupId = groupId;
        return this.getOnlyActiveUsers(args, logger);
      })
      .then((resolvedUsernameList: string[]): Promise<SharingResult> => {
        if (this.verbose) {
          logger.logToStderr(`Start adding Active user/s to SharePoint Group ${args.options.groupId ? args.options.groupId : args.options.groupName}`);
        }

        const data: any = {
          url: args.options.webUrl,
          peoplePickerInput: this.getFormattedUserList(resolvedUsernameList),
          roleValue: `group:${this.groupId}`
        };

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/SP.Web.ShareObject`,
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=verbose'
          },
          data: data,
          responseType: 'json'
        };

        return request.post<SharingResult>(requestOptions);
      })
      .then((sharingResult: SharingResult): void => {
        if (sharingResult.ErrorMessage !== null) {
          return cb(new CommandError(sharingResult.ErrorMessage));
        }

        logger.log(sharingResult.UsersAddedToGroup);

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getGroupId(args: CommandArgs): Promise<number> {
    if (args.options.groupId) {
      return Promise.resolve(args.options.groupId);
    }

    const requestUrl: string = `${args.options.webUrl}/_api/web/sitegroups/GetByName('${encodeURIComponent(args.options.groupName as string)}')`;

    const requestOptions: any = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request
      .get<{ Id: number }>(requestOptions)
      .then(response => {
        const groupId: number | undefined = response.Id;

        if (!groupId) {
          return Promise.reject(`The specified group not exist in the SharePoint site`);
        }

        return Promise.resolve(groupId);
      });
  }

  private getOnlyActiveUsers(args: CommandArgs, logger: Logger): Promise<string[]> {
    if (this.verbose) {
      logger.logToStderr(`Removing Users which are not active from the original list`);
    }

    const activeUsernamelist: string[] = [];
    const userInfo: string = args.options.userName ? args.options.userName : args.options.email!;

    return Promise.all(userInfo.split(",").map(singleUsername => {
      const options: AadUserGetCommandOptions = {
        output: 'json',
        debug: args.options.debug,
        verbose: args.options.verbose
      };

      if (args.options.userName) {
        options.userName = singleUsername.trim();
      }
      else {
        options.email = singleUsername.trim();
      }

      return Cli.executeCommandWithOutput(AadUserGetCommand as Command, { options: { ...options, _: [] } })
        .then((getUserGetOutput: CommandOutput): void => {
          if (this.debug) {
            logger.logToStderr(getUserGetOutput.stderr);
          }

          activeUsernamelist.push(JSON.parse(getUserGetOutput.stdout).userPrincipalName);
        }, (err: CommandErrorWithOutput) => {
          if (this.debug) {
            logger.logToStderr(err.stderr);
          }
        });
    }))
      .then((): Promise<string[]> => {
        return Promise.resolve(activeUsernamelist);
      });
  }

  private getFormattedUserList(activeUserList: string[]): any {
    const generatedPeoplePicker: any = JSON.stringify(activeUserList.map(singleUsername => {
      return { Key: singleUsername.trim() };
    }));
    return generatedPeoplePicker;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--groupId [groupId]'
      },
      {
        option: '--groupName [groupName]'
      },
      {
        option: '--userName [userName]'
      },
      {
        option: '--email [email]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
    if (isValidSharePointUrl !== true) {
      return isValidSharePointUrl;
    }

    if (!args.options.groupId && !args.options.groupName) {
      return 'Specify either groupId or groupName';
    }

    if (args.options.groupId && args.options.groupName) {
      return 'Specify either groupId or groupName but not both';
    }

    if (!args.options.userName && !args.options.email) {
      return 'Specify either userName or email';
    }

    if (args.options.userName && args.options.email) {
      return 'Specify either userName or email but not both';
    }

    if (args.options.groupId && isNaN(args.options.groupId)) {
      return `Specified groupId ${args.options.groupId} is not a number`;
    }

    return true;
  }
}

module.exports = new SpoGroupUserAddCommand();