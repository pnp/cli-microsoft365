import { z } from 'zod';
import { cli } from '../../../cli/cli.js';
import { Logger } from '../../../cli/Logger.js';
import { globalOptionsZod } from '../../../Command.js';
import AnonymousCommand from '../../base/AnonymousCommand.js';
import commands from '../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  service: z.enum(['VivaEngage']).alias('s')
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class CliConsentCommand extends AnonymousCommand {
  public get name(): string {
    return commands.CONSENT;
  }

  public get description(): string {
    return 'Consents additional permissions for the Microsoft Entra application used by the CLI for Microsoft 365';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let scope = '';
    if (args.options.service === 'VivaEngage') {
      scope = 'https://api.yammer.com/user_impersonation';
    }

    await logger.log(`To consent permissions for executing ${args.options.service} commands, navigate in your web browser to https://login.microsoftonline.com/${cli.getTenant()}/oauth2/v2.0/authorize?client_id=${cli.getClientId()}&response_type=code&scope=${encodeURIComponent(scope)}`);
  }

  public async action(logger: Logger, args: CommandArgs): Promise<void> {
    await this.initAction(args, logger);
    await this.commandAction(logger, args);
  }
}

export default new CliConsentCommand();