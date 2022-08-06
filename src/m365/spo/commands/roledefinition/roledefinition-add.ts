import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import { BasePermissions, PermissionKind } from '../../base-permissions';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  name: string;
  description?: string;
}

class SpoRoleDefinitionAddCommand extends SpoCommand {
  public get name(): string {
    return commands.ROLEDEFINITION_ADD;
  }

  public get description(): string {
    return 'Adds a new roledefinition to web';
  }

  protected get permissionsKindMap(): string[] {
    const result: string[] = [];

    for (const kind in PermissionKind) {
      if (typeof PermissionKind[kind] === 'number') {
        result.push(kind);
      }
    }
    return result;
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.rights = args.options.rights;
    telemetryProps.description = (!(!args.options.description)).toString();
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Adding role definition to ${args.options.webUrl}...`);
    }

    const description = args.options.description || '';

    const permissions: BasePermissions = new BasePermissions();
    if (args.options.rights) {
      const rights = args.options.rights.split(',');

      for (const item of rights) {
        const kind: PermissionKind = PermissionKind[(item.trim() as keyof typeof PermissionKind)];

        permissions.set(kind);
      }
    }

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web/roledefinitions`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json',
      data: {
        'BasePermissions': {
          'High': permissions.high.toString(),
          'Low': permissions.low.toString()
        },
        'Description': `${description}`,
        'Name': `${args.options.name}`
      }
    };

    request
      .post(requestOptions)
      .then(_ => cb(), (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-n, --name <name>'
      },
      {
        option: '-d, --description [description]'
      },
      {
        option: '--rights [rights]',
        autocomplete: this.permissionsKindMap
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.rights) {
      const rights = args.options.rights.split(',');

      for (const item of rights) {
        const kind: PermissionKind = PermissionKind[(item.trim() as keyof typeof PermissionKind)];

        if (!kind) {
          return `Rights option '${item}' is not recognized as valid PermissionKind choice. Please note it is case sensitive`;
        }
      }
    }

    return validation.isValidSharePointUrl(args.options.webUrl);
  }
}

module.exports = new SpoRoleDefinitionAddCommand();