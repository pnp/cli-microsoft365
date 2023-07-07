import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { validation } from '../../../../utils/validation';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appId?: string;
  appName?: string;
  objectId?: string;
}

class AadSpAddCommand extends GraphCommand {
  public get name(): string {
    return commands.SP_ADD;
  }

  public get description(): string {
    return 'Adds a service principal to a registered Azure AD app';
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
        appId: (!(!args.options.appId)).toString(),
        appName: (!(!args.options.appName)).toString(),
        objectId: (!(!args.options.objectId)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--appId [appId]'
      },
      {
        option: '--appName [appName]'
      },
      {
        option: '--objectId [objectId]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.appId && !validation.isValidGuid(args.options.appId)) {
          return `${args.options.appId} is not a valid appId GUID`;
        }

        if (args.options.objectId && !validation.isValidGuid(args.options.objectId)) {
          return `${args.options.objectId} is not a valid objectId GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['appId', 'appName', 'objectId'] });
  }

  private async getAppId(args: CommandArgs): Promise<string> {
    if (args.options.appId) {
      return args.options.appId;
    }

    let spMatchQuery: string = '';
    if (args.options.appName) {
      spMatchQuery = `displayName eq '${formatting.encodeQueryParameter(args.options.appName)}'`;
    }
    else if (args.options.objectId) {
      spMatchQuery = `id eq '${formatting.encodeQueryParameter(args.options.objectId)}'`;
    }

    const appIdRequestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/applications?$filter=${spMatchQuery}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: { appId: string; }[] }>(appIdRequestOptions);

    const spItem: { appId: string } | undefined = response.value[0];

    if (!spItem) {
      throw `The specified Azure AD app doesn't exist`;
    }

    if (response.value.length > 1) {
      throw `Multiple Azure AD apps with name ${args.options.appName} found: ${response.value.map(x => x.appId)}`;
    }

    return spItem.appId;

  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const appId = await this.getAppId(args);

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/servicePrincipals`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json;odata=nometadata'
        },
        data: {
          appId: appId
        },
        responseType: 'json'
      };

      const res = await request.post(requestOptions);

      logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new AadSpAddCommand();