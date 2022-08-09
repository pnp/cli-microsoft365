import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import { BasePermissions } from '../../base-permissions';
import commands from '../../commands';
import { RoleDefinition } from './RoleDefinition';
import { RoleType } from './RoleType';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  id: number;
}

class SpoRoleDefinitionGetCommand extends SpoCommand {
  public get name(): string {
    return commands.ROLEDEFINITION_GET;
  }

  public get description(): string {
    return 'Gets specified role definition from web by id';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-i, --id <id>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (isNaN(args.options.id)) {
          return `${args.options.id} is not a number`;
        }
    
        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Getting role definition from ${args.options.webUrl}...`);
    }

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web/roledefinitions(${args.options.id})`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get<RoleDefinition>(requestOptions)
      .then((response: RoleDefinition): void => {
        const permissions: BasePermissions = new BasePermissions();
        permissions.high = response.BasePermissions.High as number;
        permissions.low = response.BasePermissions.Low as number;
        response.BasePermissionsValue = permissions.parse();
        response.RoleTypeKindValue = RoleType[response.RoleTypeKind];

        logger.log(response);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new SpoRoleDefinitionGetCommand();