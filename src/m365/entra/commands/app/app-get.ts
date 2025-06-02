import { Application } from '@microsoft/microsoft-graph-types';
import fs from 'fs';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import { M365RcJson } from '../../../base/M365RcJson.js';
import commands from '../../commands.js';
import { entraApp } from '../../../../utils/entraApp.js';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  appId?: string;
  objectId?: string;
  name?: string;
  save?: boolean;
  properties?: string;
}

class EntraAppGetCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_GET;
  }

  public get description(): string {
    return 'Gets an Entra app registration';
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
        appId: typeof args.options.appId !== 'undefined',
        objectId: typeof args.options.objectId !== 'undefined',
        name: typeof args.options.name !== 'undefined',
        properties: typeof args.options.properties !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '--appId [appId]' },
      { option: '--objectId [objectId]' },
      { option: '--name [name]' },
      { option: '--save' },
      { option: '-p, --properties [properties]' }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.appId && !validation.isValidGuid(args.options.appId as string)) {
          return `${args.options.appId} is not a valid GUID`;
        }

        if (args.options.objectId && !validation.isValidGuid(args.options.objectId as string)) {
          return `${args.options.objectId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['appId', 'objectId', 'name'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const appInfo = await this.getAppInfo(args, logger);
      const res = await this.saveAppInfo(args, appInfo, logger);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getAppInfo(args: CommandArgs, logger: Logger): Promise<Application> {
    const { objectId, appId, name } = args.options;
    const properties = args.options.properties?.split(',');

    if (this.verbose) {
      await logger.logToStderr(`Retrieving information about Microsoft Entra app ${objectId ? objectId : (appId ? appId : name)}...`);
    }

    if (objectId) {
      return await entraApp.getAppRegistrationByObjectId(objectId, properties);
    }
    else if (appId) {
      return await entraApp.getAppRegistrationByAppId(appId, properties);
    }
    else {
      return await entraApp.getAppRegistrationByAppName(name as string, properties);
    }
  }

  private async saveAppInfo(args: CommandArgs, appInfo: Application, logger: Logger): Promise<Application> {
    if (!args.options.save) {
      return appInfo;
    }

    const filePath: string = '.m365rc.json';

    if (this.verbose) {
      await logger.logToStderr(`Saving Microsoft Entra app registration information to the ${filePath} file...`);
    }

    let m365rc: M365RcJson = {};
    if (fs.existsSync(filePath)) {
      if (this.debug) {
        await logger.logToStderr(`Reading existing ${filePath}...`);
      }

      try {
        const fileContents: string = fs.readFileSync(filePath, 'utf8');
        if (fileContents) {
          m365rc = JSON.parse(fileContents);
        }
      }
      catch (e) {
        await logger.logToStderr(`Error reading ${filePath}: ${e}. Please add app info to ${filePath} manually.`);
        return Promise.resolve(appInfo);
      }
    }

    if (!m365rc.apps) {
      m365rc.apps = [];
    }

    if (!m365rc.apps.some(a => a.appId === appInfo.appId)) {
      m365rc.apps.push({
        appId: appInfo.appId as string,
        name: appInfo.displayName as string
      });

      try {
        fs.writeFileSync(filePath, JSON.stringify(m365rc, null, 2));
      }
      catch (e) {
        await logger.logToStderr(`Error writing ${filePath}: ${e}. Please add app info to ${filePath} manually.`);
      }
    }

    return appInfo;
  }
}

export default new EntraAppGetCommand();