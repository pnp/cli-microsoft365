import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appId?: string;
  objectId?: string;
  name?: string;
  uri?: string;
}

class AadAppSetCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_SET;
  }

  public get description(): string {
    return 'Updates Azure AD app registration';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.appId = typeof args.options.appId !== 'undefined';
    telemetryProps.objectId = typeof args.options.objectId !== 'undefined';
    telemetryProps.name = typeof args.options.name !== 'undefined';
    telemetryProps.uri = typeof args.options.uri !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getAppObjectId(args, logger)
      .then(objectId => this.configureUri(args, objectId, logger))
      .then(_ => cb(), (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }

  private getAppObjectId(args: CommandArgs, logger: Logger): Promise<string> {
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

        return Promise.reject(`Multiple Azure AD application registration with name ${name} found. Please disambiguate (app object IDs): ${res.value.map(a => a.id).join(', ')}`);
      });
  }

  private configureUri(args: CommandArgs, objectId: string, logger: Logger): Promise<void> {
    if (!args.options.uri) {
      return Promise.resolve();
    }

    if (this.verbose) {
      logger.logToStderr(`Configuring Azure AD application ID URI...`);
    }

    const applicationInfo: any = {
      identifierUris: [args.options.uri]
    };

    const requestOptions: any = {
      url: `${this.resource}/v1.0/myorganization/applications/${objectId}`,
      headers: {
        'content-type': 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: applicationInfo
    };

    return request.patch(requestOptions);
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      { option: '--appId [appId]' },
      { option: '--objectId [objectId]' },
      { option: '-n, --name [name]' },
      { option: '-u, --uri [uri]' }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.appId &&
      !args.options.objectId &&
      !args.options.name) {
      return 'Specify either appId, objectId or name';
    }

    if ((args.options.appId && args.options.objectId) ||
      (args.options.appId && args.options.name) ||
      (args.options.objectId && args.options.name)) {
      return 'Specify either appId, objectId or name but not both';
    }

    return true;
  }
}

module.exports = new AadAppSetCommand();