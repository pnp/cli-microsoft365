import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  grantId: string;
  scope: string;
}

class AadOAuth2GrantSetCommand extends GraphCommand {
  public get name(): string {
    return commands.OAUTH2GRANT_SET;
  }

  public get description(): string {
    return 'Update OAuth2 permissions for the service principal';
  }

  constructor() {
    super();

    this.#initOptions();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --grantId <grantId>'
      },
      {
        option: '-s, --scope <scope>'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Updating OAuth2 permissions...`);
    }

    try {
      const requestOptions: any = {
        url: `${this.resource}/v1.0/oauth2PermissionGrants/${formatting.encodeQueryParameter(args.options.grantId)}`,
        headers: {
          'content-type': 'application/json'
        },
        responseType: 'json',
        data: {
          "scope": args.options.scope
        }
      };

      await request.patch(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new AadOAuth2GrantSetCommand();