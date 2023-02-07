import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { validation } from '../../../../utils/validation';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { odata } from '../../../../utils/odata';
import { formatting } from '../../../../utils/formatting';
import { User } from '@microsoft/microsoft-graph-types';


interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  rosterId: string;
  userId?: string;
  userName?: string;
  confirm?: boolean;
}

class PlannerRosterMemberRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.ROSTER_MEMBER_REMOVE;
  }

  public get description(): string {
    return 'Removes a member from a Microsoft Planner Roster';
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
        userName: typeof args.options.userName !== 'undefined',
        confirm: !!args.options.confirm
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--rosterId <rosterId>'
      },
      {
        option: '--userId [userId]'
      },
      {
        option: '--userName [userName]'
      },
      {
        option: '--confirm'
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['userId', 'userName'] }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr('Removing a member from a Microsoft Planner Roster');
    }

    if (args.options.confirm) {
      await this.removeRosterMember(args, logger);
    }
    else {
      const rosterMembers = await this.getRosterMembers(args);
      let message = '';
      if (rosterMembers === 1) {
        message = `Are you sure you want to remove the last member from the roster '${args.options.rosterId}'?`;
      }
      else {
        message = `Are you sure you want to remove member '${args.options.userId || args.options.userName}'?`;
      }
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: message
      });

      if (result.continue) {
        await this.removeRosterMember(args, logger);
      }
    }
  }

  private async getUserId(logger: Logger, args: CommandArgs): Promise<string> {
    if (this.verbose) {
      logger.logToStderr("Getting the user ID");
    }

    if (args.options.userId) {
      return args.options.userId;
    }

    const requestUrl: string = `${this.resource}/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(args.options.userName as string)}'`;

    const requestOptions: any = {
      url: requestUrl,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const res = await request.get<{ value: User[] }>(requestOptions);

    if (res.value.length === 0) {
      throw `The specified user with user name ${args.options.userName} does not exist`;
    }

    return res.value[0].id!;
  }

  private async removeRosterMember(args: CommandArgs, logger: Logger): Promise<void> {
    try {
      const userId = await this.getUserId(logger, args);

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/beta/planner/rosters/${args.options.rosterId}/members/${userId}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      await request.delete(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getRosterMembers(args: CommandArgs): Promise<number> {
    const response = await odata.getAllItems(`${this.resource}/beta/planner/rosters/${args.options.rosterId}/members`);
    return response.length;
  }
}

module.exports = new PlannerRosterMemberRemoveCommand();
