import { TeamsAppDefinition, TeamsAppInstallation } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { odata } from '../../../../utils/odata';
import { validation } from '../../../../utils/validation';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userId: string;
  userName: string;
}

class TeamsUserAppListCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_APP_LIST;
  }

  public get description(): string {
    return 'List the apps installed in the personal scope of the specified user';
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
        userId: typeof args.options.userId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--userId [userId]'
      },
      {
        option: '--userName [userName]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid GUID`;
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `${args.options.userName} is not a valid userName`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['userId', 'userName'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const userId: string = (await this.getUserId(args)).value;
      const endpoint: string = `${this.resource}/v1.0/users/${formatting.encodeQueryParameter(userId)}/teamwork/installedApps?$expand=teamsAppDefinition`;

      const items = await odata.getAllItems<TeamsAppInstallation>(endpoint);
      items.forEach(i => {
        const userAppId: string = Buffer.from(i.id as string, 'base64').toString('ascii');
        const appId: string = userAppId.substr(userAppId.indexOf("##") + 2, userAppId.length - userId.length - 2);
        (i as any).appId = appId;
      });

      if (args.options.output === 'json') {
        logger.log(items);
      }
      else {
        logger.log(items.map(i => {
          return {
            id: i.id,
            appId: (i as any).appId,
            displayName: (i.teamsAppDefinition as TeamsAppDefinition).displayName,
            version: (i.teamsAppDefinition as TeamsAppDefinition).version
          };
        }));
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getUserId(args: CommandArgs): Promise<{ value: string }> {
    if (args.options.userId) {
      return Promise.resolve({ value: args.options.userId });
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/users/${formatting.encodeQueryParameter(args.options.userName)}/id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.get<{ value: string; }>(requestOptions);
  }
}

module.exports = new TeamsUserAppListCommand();