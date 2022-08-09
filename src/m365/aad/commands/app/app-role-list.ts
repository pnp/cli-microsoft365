import { AppRole } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { odata } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appId?: string;
  appObjectId?: string;
  appName?: string;
}

class AadAppRoleListCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_ROLE_LIST;
  }

  public get description(): string {
    return 'Gets Azure AD app registration roles';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        appId: typeof args.options.appId !== 'undefined',
        appObjectId: typeof args.options.appObjectId !== 'undefined',
        appName: typeof args.options.appName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '--appId [appId]' },
      { option: '--appObjectId [appObjectId]' },
      { option: '--appName [appName]' }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const { appId, appObjectId, appName } = args.options;

        if ((appId && appObjectId) ||
          (appId && appName) ||
          (appObjectId && appName)) {
          return `Specify either appId, appObjectId or appName but not multiple`;
        }
    
        if (!appId && !appObjectId && !appName) {
          return `Specify either appId, appObjectId or appName`;
        }
    
        return true;
      }
    );
  }

  public defaultProperties(): string[] | undefined {
    return ['displayName', 'description', 'id'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getAppObjectId(args, logger)
      .then(objectId => odata.getAllItems<AppRole>(`${this.resource}/v1.0/myorganization/applications/${objectId}/appRoles`))
      .then(appRoles => {
        logger.log(appRoles);
        cb();
      }, rawRes => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }

  private getAppObjectId(args: CommandArgs, logger: Logger): Promise<string> {
    if (args.options.appObjectId) {
      return Promise.resolve(args.options.appObjectId);
    }

    const { appId, appName } = args.options;

    if (this.verbose) {
      logger.logToStderr(`Retrieving information about Azure AD app ${appId ? appId : appName}...`);
    }

    const filter: string = appId ?
      `appId eq '${encodeURIComponent(appId)}'` :
      `displayName eq '${encodeURIComponent(appName as string)}'`;

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
          const applicationIdentifier = appId ? `ID ${appId}` : `name ${appName}`;
          return Promise.reject(`No Azure AD application registration with ${applicationIdentifier} found`);
        }

        return Promise.reject(`Multiple Azure AD application registration with name ${appName} found. Please disambiguate (app object IDs): ${res.value.map(a => a.id).join(', ')}`);
      });
  }
}

module.exports = new AadAppRoleListCommand();