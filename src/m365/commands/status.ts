import auth, { AuthType, CloudType } from '../../Auth.js';
import { Logger } from '../../cli/Logger.js';
import Command, { CommandArgs, CommandError } from '../../Command.js';
import { accessToken } from '../../utils/accessToken.js';
import commands from './commands.js';

class StatusCommand extends Command {
  public get name(): string {
    return commands.STATUS;
  }

  public get description(): string {
    return 'Shows Microsoft 365 login status';
  }

  public async commandAction(logger: Logger): Promise<void> {
    if (auth.service.connected) {
      try {
        await auth.ensureAccessToken(auth.defaultResource, logger, this.debug);
      }
      catch (err: any) {
        if (this.debug) {
          await logger.logToStderr(err);
        }

        auth.service.logout();
        throw new CommandError(`Your login has expired. Sign in again to continue. ${err.message}`);
      }

      if (this.debug) {
        await logger.logToStderr({
          connectedAs: accessToken.getUserNameFromAccessToken(auth.service.accessTokens[auth.defaultResource].accessToken),
          authType: AuthType[auth.service.authType],
          appId: auth.service.appId,
          appTenant: auth.service.tenant,
          accessTokens: JSON.stringify(auth.service.accessTokens, null, 2),
          cloudType: CloudType[auth.service.cloudType]
        });
      }
      else {
        await logger.log({
          connectedAs: accessToken.getUserNameFromAccessToken(auth.service.accessTokens[auth.defaultResource].accessToken),
          authType: AuthType[auth.service.authType],
          appId: auth.service.appId,
          appTenant: auth.service.tenant,
          cloudType: CloudType[auth.service.cloudType]
        });
      }
    }
    else {
      if (this.verbose) {
        await logger.logToStderr('Logged out from Microsoft 365');
      }
      else {
        await logger.log('Logged out');
      }
    }
  }

  public async action(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      await auth.restoreAuth();
    }
    catch (error: any) {
      throw new CommandError(error);
    }

    this.initAction(args, logger);
    await this.commandAction(logger);
  }
}

export default new StatusCommand();