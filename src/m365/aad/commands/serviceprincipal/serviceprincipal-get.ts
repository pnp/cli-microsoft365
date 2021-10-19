import { ServicePrincipal } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  appId?: string;
  name?: string;
}

class AadServicePrincipalGetCommand extends GraphCommand {
  public get name(): string {
    return commands.SERVICEPRINCIPAL_GET;
  }

  public get description(): string {
    return 'Retrieves a service principal from Azure AD directory';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = typeof args.options.id !== 'undefined';
    telemetryProps.appId = typeof args.options.appId !== 'undefined';
    telemetryProps.name = typeof args.options.name !== 'undefined';
    return telemetryProps;
  }

  private getServicePrincipalId(args: CommandArgs): Promise<string> {
    if (args.options.id) {
      return Promise.resolve(args.options.id);
    }

    let requestURL = '';
    if (args.options.appId) {
      requestURL = `${this.resource}/v1.0/serviceprincipals?$filter=appId eq '${encodeURIComponent(args.options.appId)}'&$select=id`;
    }
    else {
      requestURL = `${this.resource}/v1.0/serviceprincipals?$filter=displayName eq '${encodeURIComponent(args.options.name as string)}'&$select=id`;
    }

    const requestOptions: any = {
      url: requestURL,
      headers: {
        accept: 'application/json;odata.metadata=none',
        ConsistencyLevel: 'eventual'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: [ServicePrincipal] }>(requestOptions)
      .then(response => {
        const servicePrincipalItem: ServicePrincipal | undefined = response.value[0];

        if (!servicePrincipalItem) {
          return Promise.reject(`The specified service principal doesn't exist in Azure AD`);
        }

        if (response.value.length > 1 && args.options.name) {
          return Promise.reject(`Multiple service principals with name ${args.options.name} found: ${response.value.map(x => x.id)}`);
        }
        if (response.value.length > 1 && args.options.appId) {
          return Promise.reject(`Multiple service principals with appId ${args.options.appId} found: ${response.value.map(x => x.id)}`);
        }

        return Promise.resolve(servicePrincipalItem.id as string);
      });
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getServicePrincipalId(args)
      .then((servicePrincipalId: string): Promise<ServicePrincipal> => {
        const requestOptions: any = {
          url: `${this.resource}/v1.0/serviceprincipals/${encodeURIComponent(servicePrincipalId)}`,
          headers: {
            accept: 'application/json;odata.metadata=none',
            ConsistencyLevel: 'eventual'
          },
          responseType: 'json'
        };

        return request.get<ServicePrincipal>(requestOptions);
      })
      .then((res: ServicePrincipal): void => {
        logger.log(res);
        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id [id]'
      },
      {
        option: '--appId [appId]'
      },
      {
        option: '--n, --name [name]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.id && args.options.appId && args.options.name) {
      return 'Specify either id, appId or name, but not all.';
    }

    if (args.options.id && args.options.appId) {
      return 'Specify either id or appId, but not both.';
    }

    if (args.options.appId && args.options.name) {
      return 'Specify either appId or name, but not both.';
    }

    if (args.options.name && args.options.id) {
      return 'Specify either id or name, but not both.';
    }

    if (!args.options.id && !args.options.appId && !args.options.name) {
      return 'Specify id, appId or name, one is required';
    }

    if (args.options.id && !Utils.isValidGuid(args.options.id as string)) {
      return `${args.options.id} is not a valid GUID`;
    }

    if (args.options.appId && !Utils.isValidGuid(args.options.appId as string)) {
      return `${args.options.appId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new AadServicePrincipalGetCommand();
