import { AppRole } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { entraApp } from '../../../../utils/entraApp.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appId?: string;
  appObjectId?: string;
  appName?: string;
}

class EntraAppRoleListCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_ROLE_LIST;
  }

  public get description(): string {
    return 'Gets Entra app registration roles';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initOptionSets();
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

  #initOptionSets(): void {
    this.optionSets.push({ options: ['appId', 'appObjectId', 'appName'] });
  }

  public defaultProperties(): string[] | undefined {
    return ['displayName', 'description', 'id'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const objectId = await this.getAppObjectId(args, logger);
      const appRoles = await odata.getAllItems<AppRole>(`${this.resource}/v1.0/myorganization/applications/${objectId}/appRoles`);
      await logger.log(appRoles);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getAppObjectId(args: CommandArgs, logger: Logger): Promise<string> {
    if (args.options.appObjectId) {
      return args.options.appObjectId;
    }

    const { appId, appName } = args.options;

    if (this.verbose) {
      await logger.logToStderr(`Retrieving information about Microsoft Entra app ${appId ? appId : appName}...`);
    }

    if (appId) {
      const app = await entraApp.getAppRegistrationByAppId(appId, ["id"]);
      return app.id!;
    }
    else {
      const app = await entraApp.getAppRegistrationByAppName(appName!, ["id"]);
      return app.id!;
    }
  }
}

export default new EntraAppRoleListCommand();