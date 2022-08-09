import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  clientId: string;
  resourceId: string;
  scope: string;
}

class AadOAuth2GrantAddCommand extends GraphCommand {
  public get name(): string {
    return commands.OAUTH2GRANT_ADD;
  }

  public get description(): string {
    return 'Grant the specified service principal OAuth2 permissions to the specified resource';
  }

  constructor() {
    super();
  
    this.#initOptions();
    this.#initValidators();
  }
  
  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --clientId <clientId>'
      },
      {
        option: '-r, --resourceId <resourceId>'
      },
      {
        option: '-s, --scope <scope>'
      }
    );
  }
  
  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.clientId)) {
          return `${args.options.clientId} is not a valid GUID`;
        }
    
        if (!validation.isValidGuid(args.options.resourceId)) {
          return `${args.options.resourceId} is not a valid GUID`;
        }
    
        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Granting the service principal specified permissions...`);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/oauth2PermissionGrants`,
      headers: {
        'content-type': 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: {
        "clientId": args.options.clientId,
        "consentType": "AllPrincipals",
        "principalId": null,
        "resourceId": args.options.resourceId,
        "scope": args.options.scope
      }
    };

    request
      .post<void>(requestOptions)
      .then(_ => cb(), (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }
}

module.exports = new AadOAuth2GrantAddCommand();
