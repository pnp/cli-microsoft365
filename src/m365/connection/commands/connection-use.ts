import { z } from 'zod';
import { Logger } from '../../../cli/Logger.js';
import auth, { Connection } from '../../../Auth.js';
import commands from '../commands.js';
import Command, { CommandError, globalOptionsZod } from '../../../Command.js';
import { formatting } from '../../../utils/formatting.js';
import { cli } from '../../../cli/cli.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  name: z.string().optional().alias('n')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class ConnectionUseCommand extends Command {
  public get name(): string {
    return commands.USE;
  }

  public get description(): string {
    return 'Activate the specified Microsoft 365 tenant connection';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let connection: Connection;
    if (args.options.name) {
      connection = await auth.getConnection(args.options.name);
    }
    else {
      const connections = await auth.getAllConnections();
      connections.sort((a, b) => a.name!.localeCompare(b.name!));
      const keyValuePair = formatting.convertArrayToHashTable('name', connections);
      connection = await cli.handleMultipleResultsFound<Connection>('Please select the connection you want to activate.', keyValuePair);
    }

    if (this.verbose) {
      await logger.logToStderr(`Switching to connection '${connection.identityName}', appId: ${connection.appId}, tenantId: ${connection.identityTenantId}...`);
    }

    await auth.switchToConnection(connection);

    const details = auth.getConnectionDetails(auth.connection);

    if (this.debug) {
      (details as any).accessTokens = JSON.stringify(auth.connection.accessTokens, null, 2);
    }

    await logger.log(details);
  }

  public async action(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      await auth.restoreAuth();
    }
    catch (error: any) {
      throw new CommandError(error);
    }

    await this.initAction(args, logger);
    await this.commandAction(logger, args);
  }
}

export default new ConnectionUseCommand();