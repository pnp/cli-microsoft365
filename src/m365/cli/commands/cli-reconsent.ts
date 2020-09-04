import commands from '../commands';
import GlobalOptions from '../../../GlobalOptions';
import AnonymousCommand from '../../base/AnonymousCommand';
import config from '../../../config';

const vorpal: Vorpal = require('../../../vorpal-init');

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

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    cmd.log(`To re-consent the PnP Microsoft 365 Management Shell Azure AD application navigate in your web browser to https://login.microsoftonline.com/${config.tenant}/oauth2/authorize?client_id=${config.cliAadAppId}&response_type=code&prompt=admin_consent`);
    cb();
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(commands.RECONSENT).helpInformation());
    log(
      `  Examples:
  
    Get the URL to open in the browser to re-consent CLI for Microsoft 365 permissions
      ${this.getCommandName()}

  More information:

    Re-consent the PnP Microsoft 365 Management Shell Azure AD application
      https://pnp.github.io/cli-microsoft365/user-guide/connecting-office-365/#re-consent-the-pnp-office-365-management-shell-azure-ad-application
`);
  }
}

module.exports = new CliReconsentCommand();