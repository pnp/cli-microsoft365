import * as assert from 'assert';
import O365MgmtCommand from './O365MgmtCommand';

class MockCommand extends O365MgmtCommand {
  public get name(): string {
    return 'mock';
  }

  public get description(): string {
    return 'Mock command';
  }

  public async commandAction(): Promise<void> {
  }

  public commandHelp(): void {
  }
}

describe('O365MgmtCommand', () => {
  it('defines correct resource', () => {
    const cmd = new MockCommand();
    assert.strictEqual((cmd as any).resource, 'https://manage.office.com');
  });
});
