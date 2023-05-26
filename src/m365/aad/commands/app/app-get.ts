import { Application } from '@microsoft/microsoft-graph-types';
import * as fs from 'fs';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { validation } from '../../../../utils/validation';
import GraphCommand from '../../../base/GraphCommand';
import { M365RcJson } from '../../../base/M365RcJson';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  appId?: string;
  objectId?: string;
  name?: string;
  save?: boolean;
}

class AadAppGetCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_GET;
  }

  public get description(): string {
    return 'Gets an Azure AD app registration';
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
        name: typeof args.options.name !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '--appId [appId]' },
      { option: '--objectId [objectId]' },
      { option: '--name [name]' },
      { option: '--save' }
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
      const appObjectId = await this.getAppObjectId(args);
      const appInfo = await this.getAppInfo(appObjectId);
      const res = await this.saveAppInfo(args, appInfo, logger);
      logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getAppObjectId(args: CommandArgs): Promise<string> {
    if (args.options.objectId) {
      return args.options.objectId;
    }

    const { appId, name } = args.options;

    const filter: string = appId ?
      `appId eq '${formatting.encodeQueryParameter(appId)}'` :
      `displayName eq '${formatting.encodeQueryParameter(name as string)}'`;

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
      const applicationIdentifier = appId ? `ID ${appId}` : `name ${name}`;
      throw `No Azure AD application registration with ${applicationIdentifier} found`;
    }

    throw `Multiple Azure AD application registration with name ${name} found. Please disambiguate (app object IDs): ${res.value.map(a => a.id).join(', ')}`;
  }

  private async getAppInfo(appObjectId: string): Promise<Application> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/myorganization/applications/${appObjectId}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return await request.get<Application>(requestOptions);
  }

  private async saveAppInfo(args: CommandArgs, appInfo: Application, logger: Logger): Promise<Application> {
    if (!args.options.save) {
      return appInfo;
    }

    const filePath: string = '.m365rc.json';

    if (this.verbose) {
      logger.logToStderr(`Saving Azure AD app registration information to the ${filePath} file...`);
    }

    let m365rc: M365RcJson = {};
    if (fs.existsSync(filePath)) {
      if (this.debug) {
        logger.logToStderr(`Reading existing ${filePath}...`);
      }

      try {
        const fileContents: string = fs.readFileSync(filePath, 'utf8');
        if (fileContents) {
          m365rc = JSON.parse(fileContents);
        }
      }
      catch (e) {
        logger.logToStderr(`Error reading ${filePath}: ${e}. Please add app info to ${filePath} manually.`);
        return appInfo;
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
        logger.logToStderr(`Error writing ${filePath}: ${e}. Please add app info to ${filePath} manually.`);
      }
    }

    return appInfo;
  }
}

module.exports = new AadAppGetCommand();