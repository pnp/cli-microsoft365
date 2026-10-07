import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import VivaEngageCommand from '../../../base/VivaEngageCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  groupId: z.coerce.number().optional(),
  letter: z.string().optional(),
  reverse: z.boolean().optional(),
  limit: z.coerce.number().optional(),
  sortBy: z.enum(['messages', 'followers']).optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class VivaEngageUserListCommand extends VivaEngageCommand {
  protected items!: any[];

  public get name(): string {
    return commands.ENGAGE_USER_LIST;
  }

  public get description(): string {
    return 'Returns users from the current network';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'full_name', 'email'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => !opts.letter || /^(?!\d)[a-zA-Z]+$/i.test(opts.letter), {
        message: "Value of 'letter' is invalid. Only characters within the ranges [A - Z], [a - z] are allowed.",
        params: { customCode: 'required' }
      })
      .refine(opts => !opts.letter || opts.letter.length === 1, {
        message: "Only one char as value of 'letter' accepted.",
        params: { customCode: 'required' }
      });
  }

  private getAllItems(logger: Logger, args: CommandArgs, page: number): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      if (page === 1) {
        this.items = [];
      }

      let endPoint = `${this.resource}/v1/users.json`;

      if (args.options.groupId !== undefined) {
        endPoint = `${this.resource}/v1/users/in_group/${args.options.groupId}.json`;
      }

      endPoint += `?page=${page}`;
      if (args.options.reverse !== undefined) {
        endPoint += `&reverse=true`;
      }
      if (args.options.sortBy !== undefined) {
        endPoint += `&sort_by=${args.options.sortBy}`;
      }
      if (args.options.letter !== undefined) {
        endPoint += `&letter=${args.options.letter}`;
      }

      const requestOptions: CliRequestOptions = {
        url: endPoint,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      request
        .get(requestOptions)
        .then((res: any): void => {
          let userOutput = res;
          // groups user retrieval returns a user array containing the user objects
          if (res.users) {
            userOutput = res.users;
          }

          this.items = this.items.concat(userOutput);

          // this is executed once at the end if the limit operation has been executed
          // we need to return the array of the desired size. The API does not provide such a feature
          if (args.options.limit !== undefined && this.items.length > args.options.limit) {
            this.items = this.items.slice(0, args.options.limit);
            resolve();
          }
          else {
            // if the groups endpoint is used, the more_available will tell if a new retrieval is required
            // if the user endpoint is used, we need to page by 50 items (hardcoded)
            if (res.more_available === true || this.items.length % 50 === 0) {
              this.getAllItems(logger, args, ++page)
                .then((): void => {
                  resolve();
                }, (err: any): void => {
                  reject(err);
                });
            }
            else {
              resolve();
            }
          }
        }, (err: any): void => {
          reject(err);
        });
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    this.items = []; // this will reset the items array in interactive mode

    try {
      await this.getAllItems(logger, args, 1);
      await logger.log(this.items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new VivaEngageUserListCommand();