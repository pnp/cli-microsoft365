import { User } from '@microsoft/microsoft-graph-types';
import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { odata } from '../../../../utils/odata';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  confirm?: boolean;
}

class AadUserRecycleBinItemClearCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_RECYCLEBINITEM_CLEAR;
  }

  public get description(): string {
    return 'Removes all users from the tenant recycle bin';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        confirm: !!args.options.confirm
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--confirm'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const clearRecycleBinUsers: () => Promise<void> = async (): Promise<void> => {
      try {
        const users = await odata.getAllItems<User>(`${this.resource}/v1.0/directory/deletedItems/microsoft.graph.user?$select=id`);
        if (this.verbose) {
          logger.logToStderr(`Amount of users to permanently delete: ${users.length}`);
        }
        const batchRequests = users.map((user, index) => {
          return {
            id: index,
            method: 'DELETE',
            url: `/directory/deletedItems/${user.id}`
          };
        });
        for (let i = 0; i < batchRequests.length; i += 20) {
          const batchRequestChunk = batchRequests.slice(i, i + 20);
          if (this.verbose) {
            logger.logToStderr(`Deleting users: ${i + batchRequestChunk.length}/${users.length}`);
          }

          const requestOptions: CliRequestOptions = {
            url: `${this.resource}/v1.0/$batch`,
            headers: {
              accept: 'application/json',
              'content-type': 'application/json'
            },
            responseType: 'json',
            data: {
              requests: batchRequestChunk
            }
          };
          await request.post(requestOptions);
        }
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.confirm) {
      await clearRecycleBinUsers();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: 'Are you sure you want to permanently delete all deleted users?'
      });

      if (result.continue) {
        await clearRecycleBinUsers();
      }
    }
  }
}

module.exports = new AadUserRecycleBinItemClearCommand();
