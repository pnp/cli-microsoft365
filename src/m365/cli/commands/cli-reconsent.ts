import { Logger } from '../../../cli';
import config from '../../../config';
import GlobalOptions from '../../../GlobalOptions';
import AnonymousCommand from '../../base/AnonymousCommand';
import commands from '../commands';

interface CommandArgs {
  options: GlobalOptions;
}

class CliReconsentCommand extends AnonymousCommand {
  public get name(): string {
    return commands.RECONSENT;
  }

  public get description(): string {
    return 'Returns Azure AD URL to open in the browser to re-consent CLI for Microsoft 365 permissions';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    logger.log(`To re-consent the PnP Microsoft 365 Management Shell Azure AD application navigate in your web browser to https://login.microsoftonline.com/${config.tenant}/oauth2/authorize?client_id=${config.cliAadAppId}&response_type=code&prompt=admin_consent`);
    cb();
  }
}

module.exports = new CliReconsentCommand();