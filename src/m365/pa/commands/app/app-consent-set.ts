import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import PowerAppsCommand from '../../../base/PowerAppsCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  environmentName: z.string().alias('e'),
  name: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value is not a valid GUID.'
    })
    .alias('n'),
  bypass: z.boolean().alias('b'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PaAppConsentSetCommand extends PowerAppsCommand {
  public get name(): string {
    return commands.APP_CONSENT_SET;
  }

  public get description(): string {
    return 'Configures if users can bypass the API Consent window for the selected canvas app';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Setting the bypass consent for the Microsoft Power App ${args.options.name}... to ${args.options.bypass}`);
    }

    if (args.options.force) {
      await this.consentPaApp(args);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you bypass the consent for the Microsoft Power App ${args.options.name} to ${args.options.bypass}?` });

      if (result) {
        await this.consentPaApp(args);
      }
    }
  }

  private async consentPaApp(args: CommandArgs): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/providers/Microsoft.PowerApps/scopes/admin/environments/${args.options.environmentName}/apps/${args.options.name}/setPowerAppConnectionDirectConsentBypass?api-version=2021-02-01`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: {
        bypassconsent: args.options.bypass
      }
    };

    try {
      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PaAppConsentSetCommand();