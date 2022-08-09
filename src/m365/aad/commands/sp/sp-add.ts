import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
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
        let optionsSpecified: number = 0;
        optionsSpecified += args.options.appId ? 1 : 0;
        optionsSpecified += args.options.appName ? 1 : 0;
        optionsSpecified += args.options.objectId ? 1 : 0;
    
        if (optionsSpecified !== 1) {
          return 'Specify either appId, appName, or objectId';
        }
    
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

  private getAppId(args: CommandArgs): Promise<string> {
    if (args.options.appId) {
      return Promise.resolve(args.options.appId);
    }

    let spMatchQuery: string = '';
    if (args.options.appName) {
      spMatchQuery = `displayName eq '${encodeURIComponent(args.options.appName)}'`;
    }
    else if (args.options.objectId) {
      spMatchQuery = `id eq '${encodeURIComponent(args.options.objectId)}'`;
    }

    const appIdRequestOptions: any = {
      url: `${this.resource}/v1.0/applications?$filter=${spMatchQuery}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: { appId: string; }[] }>(appIdRequestOptions)
      .then(response => {
        const spItem: { appId: string } | undefined = response.value[0];

        if (!spItem) {
          return Promise.reject(`The specified Azure AD app doesn't exist`);
        }

        if (response.value.length > 1) {
          return Promise.reject(`Multiple Azure AD apps with name ${args.options.appName} found: ${response.value.map(x => x.appId)}`);
        }

        return Promise.resolve(spItem.appId);
      });
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getAppId(args)
      .then((appId: string): Promise<void> => {
        const requestOptions: any = {
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

        return request.post(requestOptions);
      })
      .then((res: any): void => {
        logger.log(res);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new AadSpAddCommand();