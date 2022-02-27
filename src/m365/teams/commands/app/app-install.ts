import { Cli, CommandOutput, Logger } from '../../../../cli';
import Command, { CommandError, CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import * as AadUserGetCommand from '../../../aad/commands/user/user-get';
import { Options as AadUserGetCommandOptions } from '../../../aad/commands/user/user-get';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appId: string;
  teamId?: string;
  userId?: string;
  userName?: string;
}

class TeamsAppInstallCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_INSTALL;
  }

  public get description(): string {
    return 'Installs a Microsoft Teams team app from the catalog in the specified team or for the specified user';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .validateUser(args, logger)
      .then(_ => {
        let url: string = `${this.resource}/v1.0`;
        if (args.options.teamId) {
          url += `/teams/${encodeURIComponent(args.options.teamId)}/installedApps`;
        }
        else {
          url += `/users/${encodeURIComponent((args.options.userId ?? args.options.userName) as string)}/teamwork/installedApps`;
        }

        const requestOptions: any = {
          url: url,
          headers: {
            'content-type': 'application/json;odata=nometadata',
            'accept': 'application/json;odata.metadata=none'
          },
          responseType: 'json',
          data: {
            'teamsApp@odata.bind': `${this.resource}/v1.0/appCatalogs/teamsApps/${args.options.appId}`
          }
        };

        return request.post(requestOptions);
      })
      .then(_ => cb(), (res: any): void => this.handleRejectedODataJsonPromise(res, logger, cb));
  }

  // we need this method, because passing an invalid user ID to the API
  // won't cause an error
  private validateUser(args: CommandArgs, logger: Logger): Promise<boolean> {
    if (!args.options.userId) {
      return Promise.resolve(true);
    }

    if (this.verbose) {
      logger.logToStderr(`Checking if user ${args.options.userId} exists...`);
    }

    const options: AadUserGetCommandOptions = {
      id: args.options.userId,
      output: 'json',
      debug: args.options.debug,
      verbose: args.options.verbose
    };

    return Cli
      .executeCommandWithOutput(AadUserGetCommand as Command, { options: { ...options, _: [] } })
      .then((res: CommandOutput) => {
        if (this.verbose) {
          logger.logToStderr(res.stderr);
        }

        return true;
      }, (err: { error: CommandError, stderr: string }) => {
        if (this.verbose) {
          logger.logToStderr(err.stderr);
        }

        return Promise.reject(`User with ID ${args.options.userId} not found. Original error: ${err.error.message}`);
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      { option: '--appId <appId>' },
      { option: '--teamId [teamId' },
      { option: '--userId [userId]' },
      { option: '--userName [userName]' }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!Utils.isValidGuid(args.options.appId)) {
      return `${args.options.appId} is not a valid GUID`;
    }

    if (!args.options.teamId &&
      !args.options.userId &&
      !args.options.userName) {
      return `Specify either teamId, userId or userName`;
    }

    if ((args.options.teamId && args.options.userId) ||
      (args.options.teamId && args.options.userName) ||
      (args.options.userId && args.options.userName)) {
      return `Specify either teamId, userId or userName but not multiple`;
    }

    if (args.options.teamId &&
      !Utils.isValidGuid(args.options.teamId)) {
      return `${args.options.teamId} is not a valid GUID`;
    }

    if (args.options.userId &&
      !Utils.isValidGuid(args.options.userId)) {
      return `${args.options.userId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new TeamsAppInstallCommand();