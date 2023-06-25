import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { formatting } from '../../../../utils/formatting';
import { odata } from '../../../../utils/odata';
import { validation } from '../../../../utils/validation';
import PowerAppsCommand from '../../../base/PowerAppsCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appName: string;
  asAdmin: boolean;
  environment?: string;
  roleName?: string;
}

class PaAppPermissionListCommand extends PowerAppsCommand {
  private allowedRoleNames = ['Owner', 'CanEdit', 'CanView'];

  public get name(): string {
    return commands.APP_PERMISSION_LIST;
  }

  public get description(): string {
    return 'Lists all permissions of a Power Apps app';
  }

  public defaultProperties(): string[] | undefined {
    return ['roleName', 'principalId', 'principalType'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        asAdmin: !!args.options.asAdmin,
        environmentName: typeof args.options.environment !== 'undefined',
        roleName: typeof args.options.roleName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--appName <appName>'
      },
      {
        option: '--asAdmin'
      },
      {
        option: '-e, --environmentName [environmentName]'
      },
      {
        option: '--roleName [roleName]',
        autocomplete: this.allowedRoleNames
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.appName)) {
          return `${args.options.appName} is not a valid GUID.`;
        }

        if (args.options.roleName && !this.allowedRoleNames.includes(args.options.roleName)) {
          return `${args.options.roleName} is not a valid roleName. Allowed values are ${this.allowedRoleNames.join('|')}`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      {
        options: ['environmentName'],
        runsWhen: (args) => args.options.asAdmin
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving permissions for app ${args.options.appName}${args.options.roleName !== undefined && ` with rolename ${args.options.roleName}`}`);
    }

    const url = `${this.resource}/providers/Microsoft.PowerApps${args.options.asAdmin ? '/scopes/admin' : ''}${args.options.environmentName ? '/environments/' + formatting.encodeQueryParameter(args.options.environmentName) : ''}/apps/${args.options.appName}/permissions?api-version=2022-11-01`;

    try {
      let permissions = await odata.getAllItems<{ principalType: string, principalId: string, roleName: string, properties: { roleName: string, principal: { id: string, type: string } } }>(url);

      if (args.options.roleName) {
        permissions = permissions.filter(permission => permission.properties.roleName === args.options.roleName);
      }

      if (permissions.length > 0) {
        permissions.forEach(permission => {
          permission.roleName = permission.properties.roleName;
          permission.principalId = permission.properties.principal.id;
          permission.principalType = permission.properties.principal.type;
        });
        logger.log(permissions);
      }
      else {
        if (this.verbose) {
          logger.logToStderr(`No permissions for app ${args.options.appName} found${args.options.roleName !== undefined ? ` with roleName ${args.options.roleName}` : ''}.`);
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new PaAppPermissionListCommand();