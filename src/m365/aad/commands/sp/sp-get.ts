import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  appId?: string;
  appDisplayName?: string;
  appObjectId?: string;
}

class AadSpGetCommand extends GraphCommand {
  public get name(): string {
    return commands.SP_GET;
  }

  public get description(): string {
    return 'Gets information about the specific service principal';
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
        appDisplayName: (!(!args.options.appDisplayName)).toString(),
        appObjectId: (!(!args.options.appObjectId)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --appId [appId]'
      },
      {
        option: '-n, --appDisplayName [appDisplayName]'
      },
      {
        option: '--appObjectId [appObjectId]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        let optionsSpecified: number = 0;
        optionsSpecified += args.options.appId ? 1 : 0;
        optionsSpecified += args.options.appDisplayName ? 1 : 0;
        optionsSpecified += args.options.appObjectId ? 1 : 0;
        if (optionsSpecified !== 1) {
          return 'Specify either appId, appObjectId or appDisplayName';
        }
    
        if (args.options.appId && !validation.isValidGuid(args.options.appId)) {
          return `${args.options.appId} is not a valid appId GUID`;
        }
    
        if (args.options.appObjectId && !validation.isValidGuid(args.options.appObjectId)) {
          return `${args.options.appObjectId} is not a valid objectId GUID`;
        }
    
        return true;
      }
    );
  }

  private getSpId(args: CommandArgs): Promise<string> {
    if (args.options.appObjectId) {
      return Promise.resolve(args.options.appObjectId);
    }

    let spMatchQuery: string = '';
    if (args.options.appDisplayName) {
      spMatchQuery = `displayName eq '${encodeURIComponent(args.options.appDisplayName)}'`;
    }
    else if (args.options.appId) {
      spMatchQuery = `appId eq '${encodeURIComponent(args.options.appId)}'`;
    }

    const idRequestOptions: any = {
      url: `${this.resource}/v1.0/servicePrincipals?$filter=${spMatchQuery}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: { id: string; }[] }>(idRequestOptions)
      .then(response => {
        const spItem: { id: string } | undefined = response.value[0];

        if (!spItem) {
          return Promise.reject(`The specified Azure AD app does not exist`);
        }

        if (response.value.length > 1) {
          return Promise.reject(`Multiple Azure AD apps with name ${args.options.appDisplayName} found: ${response.value.map(x => x.id)}`);
        }

        return Promise.resolve(spItem.id);
      });
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving service principal information...`);
    }

    this
      .getSpId(args)
      .then((id: string): Promise<void> => {
        const requestOptions: any = {
          url: `${this.resource}/v1.0/servicePrincipals/${id}`,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'content-type': 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        return request.get(requestOptions);
      })
      .then((res: any): void => {
        logger.log(res);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new AadSpGetCommand();