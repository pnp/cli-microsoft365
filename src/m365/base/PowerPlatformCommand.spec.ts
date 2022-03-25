import * as assert from 'assert';
import { Logger } from '../../cli';
import PowerPlatformCommand from './PowerPlatformCommand';

class MockCommand extends PowerPlatformCommand {
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

describe('PowerPlatformCommand', () => {
  it('returns correct bapi resource', () => {
    const command = new MockCommand();
    assert.strictEqual((command as any).resource, 'https://api.bap.microsoft.com');
  });
});
