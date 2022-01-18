import { Application } from '@microsoft/microsoft-graph-types';
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

export interface Options extends GlobalOptions {
  appId?: string;
  objectId?: string;
  name?: string;
}

class PpManagementAppAddCommand extends GraphCommand {
  public get name(): string {
    return commands.MANAGEMENTAPP_ADD;
  }

  public get description(): string {
    return 'Register management application for Power Platform';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.appId = typeof args.options.appId !== 'undefined';
    telemetryProps.objectId = typeof args.options.objectId !== 'undefined';
    telemetryProps.name = typeof args.options.name !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getAppObjectId(args)
      .then((appId: string): Promise<any> => {
        const requestOptions: any = {
          // This should be refactored once we implement a PowerPlatform base class as api.bap will differ between envs.
          url: `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/adminApplications/${appId}?api-version=2020-06-01`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        return request.put(requestOptions);
      })
      .then((res: any): void => {
        logger.log(res);
        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getAppObjectId(args: CommandArgs): Promise<string> {
    if (args.options.appId) {
      return Promise.resolve(args.options.appId);
    }

    const { objectId, name } = args.options;

    const filter: string = objectId ?
      `id eq '${encodeURIComponent(objectId)}'` :
      `displayName eq '${encodeURIComponent(name as string)}'`;

    const requestOptions: any = {
      url: `${this.resource}/v1.0/myorganization/applications?$filter=${filter}&$select=appId`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: { appId: string }[] }>(requestOptions)
      .then((res: { value: { appId: string }[] }): Promise<string> => {
        if (res.value.length === 1) {
          return Promise.resolve(res.value[0].appId);
        }

        if (res.value.length === 0) {
          const applicationIdentifier = objectId ? `ID ${objectId}` : `name ${name}`;
          return Promise.reject(`No Azure AD application registration with ${applicationIdentifier} found`);
        }

        return Promise.reject(`Multiple Azure AD application registration with name ${name} found. Please disambiguate (app IDs): ${res.value.map(a => a.appId).join(', ')}`);
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      { option: '--appId [appId]' },
      { option: '--objectId [objectId]' },
      { option: '--name [name]' }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.appId &&
      !args.options.objectId &&
      !args.options.name) {
      return 'Specify either appId, objectId, or name';
    }

    if ((args.options.appId && args.options.objectId) ||
      (args.options.appId && args.options.name) ||
      (args.options.objectId && args.options.name)) {
      return 'Specify either appId, objectId, or name but not both';
    }

    if (args.options.appId && !Utils.isValidGuid(args.options.appId as string)) {
      return `${args.options.appId} is not a valid GUID`;
    }

    if (args.options.objectId && !Utils.isValidGuid(args.options.objectId as string)) {
      return `${args.options.objectId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new PpManagementAppAddCommand();
