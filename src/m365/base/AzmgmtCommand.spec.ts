import * as assert from 'assert';
import AzmgmtCommand from './AzmgmtCommand';

class MockCommand extends AzmgmtCommand {
  public get name(): string {
    return 'mock';
  }

  public get description(): string {
    return 'Mock command';
  }

  public commandAction(): void {
  }

  public commandHelp(): void {
  }
}

describe('AzmgmtCommand', () => {
  it('defines correct resource', () => {
    const cmd = new MockCommand();
    assert.strictEqual((cmd as any).resource, 'https://management.azure.com/');
  });
});