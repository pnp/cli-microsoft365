import { AppRole } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';
import aadCommands from '../../aadCommands.js';

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

  public alias(): string[] | undefined {
    return [aadCommands.APP_ROLE_LIST, commands.APPREGISTRATION_ROLE_LIST];
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

    const filter: string = appId ?
      `appId eq '${formatting.encodeQueryParameter(appId)}'` :
      `displayName eq '${formatting.encodeQueryParameter(appName as string)}'`;

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/myorganization/applications?$filter=${filter}&$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const res = await request.get<{ value: { id: string }[] }>(requestOptions);

    if (res.value.length === 1) {
      return res.value[0].id;
    }

    if (res.value.length === 0) {
      const applicationIdentifier = appId ? `ID ${appId}` : `name ${appName}`;
      throw `No Microsoft Entra application registration with ${applicationIdentifier} found`;
    }

    const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', res.value);
    const result = await cli.handleMultipleResultsFound<{ id: string }>(`Multiple Microsoft Entra application registrations with name '${appName}' found.`, resultAsKeyValuePair);
    return result.id;
  }
}

export default new EntraAppRoleListCommand();