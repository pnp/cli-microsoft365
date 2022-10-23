import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import { BasePermissions } from '../../base-permissions';
import commands from '../../commands';
import { RoleDefinition } from './RoleDefinition';
import { RoleType } from './RoleType';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  webUrl: string;
}

class SpoRoleDefinitionListCommand extends SpoCommand {
  public get name(): string {
    return commands.ROLEDEFINITION_LIST;
  }

  public get description(): string {
    return 'Gets list of role definitions for the specified site';
  }

  public defaultProperties(): string[] | undefined {
    return ['Id', 'Name'];
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
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.webUrl)
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Getting role definitions list from ${args.options.webUrl}...`);
    }

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web/roledefinitions`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const res = await request.get<{ value: any[] }>(requestOptions);
      const response = this.setFriendlyPermissions(res.value);
      logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private setFriendlyPermissions(response: any[]): any[] {
    response.forEach((r: RoleDefinition) => {
      const permissions: BasePermissions = new BasePermissions();
      permissions.high = r.BasePermissions.High as number;
      permissions.low = r.BasePermissions.Low as number;
      r.BasePermissionsValue = permissions.parse();
      r.RoleTypeKindValue = RoleType[r.RoleTypeKind];
    });

    return response;
  }
}

module.exports = new SpoRoleDefinitionListCommand();