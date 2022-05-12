import * as assert from 'assert';
import { Logger } from '../../cli';
import PowerBICommand from './PowerBICommand';

class MockCommand extends PowerBICommand {
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

describe('PowerBICommand', () => {
  it('returns correct api resource', () => {
    const command = new MockCommand();
    assert.strictEqual((command as any).resource, 'https://api.powerbi.com');
  });
});
