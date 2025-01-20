import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import request, { CliRequestOptions } from '../../../../request.js';

const options = globalOptionsZod
  .extend({
    url: z.string().url(),
    audiences: z.string().optional(),
    vivaConnectionsDefaultStart: z.boolean().optional(),
    isInDraftMode: z.boolean().optional(),
    order: z.number().optional()
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoTenantHomeSiteAddCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_HOMESITE_ADD;
  }

  public get description(): string {
    return 'Adds a Home Site';
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
        url: args.options.url
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --url <url>'
      },
      {
        option: '--isInDraftMode [isInDraftMode]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--vivaConnectionsDefaultStart [vivaConnectionsDefaultStart]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--audiences [audiences]'
      },
      {
        option: '--order [order]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.url);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.order !== undefined && !validation.isValidPositiveInteger(args.options.order)) {
          return `${args.options.order} is not a positive integer`;
        }

        const isValidGUIDArrayResult = args.options.audiences ? validation.isValidGuidArray(args.options.audiences) : true;
        if (isValidGUIDArrayResult !== true) {
          return `The following GUIDs are invalid for the option 'ids': ${isValidGUIDArrayResult}.`;
        }
        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoAdminUrl: string = await spo.getSpoAdminUrl(logger, this.verbose);
      const requestOptions: CliRequestOptions = {
        url: `${spoAdminUrl}/_api/SPHSite/AddHomeSite`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json',
        data: {
          siteUrl: args.options.url,
          audiences: args.options.audiences?.split(','),
          vivaConnectionsDefaultStart: args.options.vivaConnectionsDefaultStart ?? true,
          isInDraftMode: args.options.isInDraftMode ?? true,
          order: args.options.order
        }
      };

      if (this.verbose) {
        await logger.logToStderr(`Adding home site with URL: ${args.options.url}...`);
      }

      const res = await request.post(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoTenantHomeSiteAddCommand();