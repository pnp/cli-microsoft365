import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  grantId: string;
}

class AadOAuth2GrantRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.OAUTH2GRANT_REMOVE;
  }

  public get description(): string {
    return 'Remove specified service principal OAuth2 permissions';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Removing OAuth2 permissions...`);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/oauth2PermissionGrants/${encodeURIComponent(args.options.grantId)}`,
      headers: {
        'accept': 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    request
      .delete(requestOptions)
      .then(_ => cb(), (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --grantId <grantId>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new AadOAuth2GrantRemoveCommand();