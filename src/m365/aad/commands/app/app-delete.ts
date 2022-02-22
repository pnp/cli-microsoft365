import { Cli, Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  appId?: string;
  objectId?: string;
  name?: string;
  confirm?: boolean;
}

class AadAppDeleteCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_DELETE;
  }

  public get description(): string {
    return 'Removes an Azure AD app registration';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.appId = typeof args.options.appId !== 'undefined';
    telemetryProps.objectId = typeof args.options.objectId !== 'undefined';
    telemetryProps.name = typeof args.options.name !== 'undefined';
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const deleteApp: () => void = (): void => {
      this
        .getObjectId(args, logger)
        .then((objectId: string): Promise<void> => {
          if (this.verbose) {
            logger.logToStderr(`Deleting Azure AD app ${objectId}...`);
          }

          const requestOptions: any = {
            url: `${this.resource}/v1.0/myorganization/applications/${objectId}`,
            headers: {
              accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
          };

          return request.delete(requestOptions);
        })
        .then(_ => cb(), (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
    };

    if (args.options.confirm) {
      deleteApp();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to delete the App?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          deleteApp();
        }
      });
    }
  }

  private getObjectId(args: CommandArgs, logger: Logger): Promise<string> {
    if (args.options.objectId) {
      return Promise.resolve(args.options.objectId);
    }

    const { appId, name } = args.options;

    if (this.verbose) {
      logger.logToStderr(`Retrieving information about Azure AD app ${appId ? appId : name}...`);
    }

    const filter: string = appId ?
      `appId eq '${encodeURIComponent(appId)}'` :
      `displayName eq '${encodeURIComponent(name as string)}'`;

    const requestOptions: any = {
      url: `${this.resource}/v1.0/myorganization/applications?$filter=${filter}&$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: { id: string }[] }>(requestOptions)
      .then((res: { value: { id: string }[] }): Promise<string> => {
        if (res.value.length === 1) {
          return Promise.resolve(res.value[0].id);
        }

        if (res.value.length === 0) {
          const applicationIdentifier = appId ? `ID ${appId}` : `name ${name}`;
          return Promise.reject(`No Azure AD application registration with ${applicationIdentifier} found`);
        }

        return Promise.reject(`Multiple Azure AD application registration with name ${name} found. Please choose one of the object IDs: ${res.value.map(a => a.id).join(', ')}`);
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      { option: '--appId [appId]' },
      { option: '--objectId [objectId]' },
      { option: '--name [name]' },
      { option: '--confirm' }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.appId &&
      !args.options.objectId &&
      !args.options.name) {
      return 'Specify either appId, objectId, or name';
    }

    if ((args.options.appId && args.options.objectId) ||
      (args.options.appId && args.options.name) ||
      (args.options.objectId && args.options.name)) {
      return 'Specify either appId, objectId, or name';
    }

    if (args.options.appId && !Utils.isValidGuid(args.options.appId as string)) {
      return `${args.options.appId} is not a valid GUID`;
    }

    if (args.options.objectId && !Utils.isValidGuid(args.options.objectId as string)) {
      return `${args.options.objectId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new AadAppDeleteCommand();