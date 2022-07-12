import * as assert from 'assert';
import { Logger } from '../../cli';
import PowerAppsCommand from './PowerAppsCommand';

class MockCommand extends PowerAppsCommand {
  public get name(): string {
    return 'mock';
  }

  public get description(): string {
    return 'Mock command';
  }

  public commandAction(logger: Logger, args: any, cb: () => void): void {
    cb();
  }

  public commandHelp(): void {
  }
}

describe('PowerAppsCommand', () => {
  it('returns correct resource', () => {
    const command = new MockCommand();
    assert.strictEqual((command as any).resource, 'https://api.powerapps.com');
  });
});