import assert from 'assert';
import fs from 'fs';
import sinon from 'sinon';
import { chili } from './chili.js';
import request from '../request.js';
import { sinonUtil } from '../utils/sinonUtil.js';
import { SelectionConfig, prompt } from '../utils/prompt.js';

describe('chili', () => {
  let consoleLogSpy: sinon.SinonStub;
  let consoleErrorSpy: sinon.SinonStub;

  before(() => {
    consoleLogSpy = sinon.stub(console, 'log').returns();
    consoleErrorSpy = sinon.stub(console, 'error').returns();
  });

  afterEach(() => {
    consoleLogSpy.resetHistory();
    consoleErrorSpy.resetHistory();

    sinonUtil.restore([
      request.post,
      prompt.forSelection,
      prompt.forInput,
      fs.existsSync
    ]);
  });

  after(() => {
    sinonUtil.restore([
      // eslint-disable-next-line no-console
      console.log,
      // eslint-disable-next-line no-console
      console.error
    ]);
  });

  it('starts a conversation using a prompt from args when specified with debug mode', async () => {
    sinon.stub(request, 'post').callsFake(async options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v1/newConversation':
          return {
            // eslint-disable-next-line camelcase
            conversation_id: 1
          };
        case 'https://api.mendable.ai/v1/mendableChat':
          if (options.data.question === 'Hello') {
            return {
              answer: {
                text: 'Hello back'
              },
              sources: []
            };
          }
          break;
        case 'https://api.mendable.ai/v1/endConversation':
          return {};
      }
      throw `Invalid request: ${options.url}`;
    });
    sinon.stub(prompt, 'forInput').resolves('Hello world');
    sinon.stub(prompt, 'forSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      if (config.message === 'Was this helpful?') {
        return 0;
      }
      else if (config.message === 'What would you like to do next?') {
        return 'end';
      }

      throw `Prompt not found for '${config.message}'`;
    });
    await assert.doesNotReject(chili.startConversation(['Hello', '--debug']));
  });

  it('starts a conversation when a prompt specified as a single arg', async () => {
    sinon.stub(request, 'post').callsFake(async options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v1/newConversation':
          return {
            // eslint-disable-next-line camelcase
            conversation_id: 1
          };
        case 'https://api.mendable.ai/v1/mendableChat':
          if (options.data.question === 'Hello world') {
            return {
              answer: {
                text: 'Hello back'
              },
              sources: []
            };
          }
          break;
        case 'https://api.mendable.ai/v1/endConversation':
          return {};
      }
      throw `Invalid request: ${options.url}`;
    });
    sinon.stub(prompt, 'forInput').resolves('Hello world');
    sinon.stub(prompt, 'forSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      if (config.message === 'Was this helpful?') {
        return 0;
      }
      else if (config.message === 'What would you like to do next?') {
        return 'end';
      }

      throw `Prompt not found for '${config.message}'`;
    });
    await assert.doesNotReject(chili.startConversation(['Hello world']));
  });

  it('starts a conversation when a prompt specified as multiple args (no quotes)', async () => {
    sinon.stub(request, 'post').callsFake(async options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v1/newConversation':
          return {
            // eslint-disable-next-line camelcase
            conversation_id: 1
          };
        case 'https://api.mendable.ai/v1/mendableChat':
          if (options.data.question === 'Hello world') {
            return {
              answer: {
                text: 'Hello back'
              },
              sources: []
            };
          }
          break;
        case 'https://api.mendable.ai/v1/endConversation':
          return {};
      }
      throw `Invalid request: ${options.url}`;
    });
    sinon.stub(prompt, 'forInput').resolves('Hello world');
    sinon.stub(prompt, 'forSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      if (config.message === 'Was this helpful?') {
        return 0;
      }
      else if (config.message === 'What would you like to do next?') {
        return 'end';
      }

      throw `Prompt not found for '${config.message}'`;
    });
    await assert.doesNotReject(chili.startConversation(['Hello', 'world']));
  });

  it('starts a conversation asking for a prompt when no prompt specified via args', async () => {
    sinon.stub(request, 'post').callsFake(async options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v1/newConversation':
          return {
            // eslint-disable-next-line camelcase
            conversation_id: 1
          };
        case 'https://api.mendable.ai/v1/mendableChat':
          if (options.data.question === 'Hello world') {
            return {
              answer: {
                text: 'Hello back'
              },
              sources: []
            };
          }
          break;
        case 'https://api.mendable.ai/v1/endConversation':
          return {};
      }
      throw `Invalid request: ${options.url}`;
    });
    sinon.stub(prompt, 'forInput').resolves('Hello world');
    sinon.stub(prompt, 'forSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      if (config.message === 'Was this helpful?') {
        return 0;
      }
      else if (config.message === 'What would you like to do next?') {
        return 'end';
      }

      throw `Prompt not found for '${config.message}'`;
    });
    await assert.doesNotReject(chili.startConversation([]));
  });

  it('uses the prompt to search in CLI docs using Mendable', async () => {
    sinon.stub(request, 'post').callsFake(async options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v1/newConversation':
          return {
            // eslint-disable-next-line camelcase
            conversation_id: 1
          };
        case 'https://api.mendable.ai/v1/mendableChat':
          if (options.data.question === 'Hello') {
            return {
              answer: {
                text: 'Hello back'
              },
              sources: []
            };
          }
          break;
        case 'https://api.mendable.ai/v1/endConversation':
          return {};
      }
      throw `Invalid request: ${options.url}`;
    });
    sinon.stub(prompt, 'forSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      if (config.message === 'Was this helpful?') {
        return 0;
      }
      else if (config.message === 'What would you like to do next?') {
        return 'end';
      }

      throw `Prompt not found for '${config.message}'`;
    });
    await assert.doesNotReject(chili.startConversation(['Hello']));
  });

  it('displays the retrieved response to the user', async () => {
    sinon.stub(request, 'post').callsFake(async options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v1/newConversation':
          return {
            // eslint-disable-next-line camelcase
            conversation_id: 1
          };
        case 'https://api.mendable.ai/v1/mendableChat':
          if (options.data.question === 'Hello') {
            return {
              answer: {
                text: 'Hello back'
              },
              sources: []
            };
          }
          break;
        case 'https://api.mendable.ai/v1/endConversation':
          return {};

      }
      throw `Invalid request: ${options.url}`;
    });
    sinon.stub(prompt, 'forSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      if (config.message === 'Was this helpful?') {
        return 0;
      }
      else if (config.message === 'What would you like to do next?') {
        return 'end';
      }

      throw `Prompt not found for '${config.message}'`;
    });
    await chili.startConversation(['Hello']);
    assert(consoleLogSpy.calledWith('Hello back'));
  });

  it('in response formats MD in terminal-friendly way', async () => {
    sinon.stub(request, 'post').callsFake(async options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v1/newConversation':
          return {
            // eslint-disable-next-line camelcase
            conversation_id: 1
          };
        case 'https://api.mendable.ai/v1/mendableChat':
          if (options.data.question === 'Hello') {
            return {
              answer: {
                text: 'Hello **back**'
              },
              sources: []
            };
          }
          break;
        case 'https://api.mendable.ai/v1/endConversation':
          return {};
      }
      throw `Invalid request: ${options.url}`;
    });
    sinon.stub(prompt, 'forSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      if (config.message === 'Was this helpful?') {
        return 0;
      }
      else if (config.message === 'What would you like to do next?') {
        return 'end';
      }

      throw `Prompt not found for '${config.message}'`;
    });
    await chili.startConversation(['Hello']);
    assert(consoleLogSpy.calledWith('Hello back'));
  });

  it('in response, shows sources', async () => {
    sinon.stub(request, 'post').callsFake(async options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v1/newConversation':
          return {
            // eslint-disable-next-line camelcase
            conversation_id: 1
          };
        case 'https://api.mendable.ai/v1/mendableChat':
          if (options.data.question === 'Hello') {
            return {
              answer: {
                text: 'Hello back'
              },
              sources: [
                {
                  link: 'https://example.com/source-1'
                }
              ]
            };
          }
          break;
        case 'https://api.mendable.ai/v1/endConversation':
          return {};
      }
      throw `Invalid request: ${options.url}`;
    });
    sinon.stub(prompt, 'forSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      if (config.message === 'Was this helpful?') {
        return 0;
      }
      else if (config.message === 'What would you like to do next?') {
        return 'end';
      }

      throw `Prompt not found for '${config.message}'`;
    });
    await chili.startConversation(['Hello']);
    assert(consoleLogSpy.calledWith('⬥ https://example.com/source-1'));
  });

  it('allows asking a follow-up question after a response', async () => {
    let questionsAsked = 0;
    sinon.stub(request, 'post').callsFake(async options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v1/newConversation':
          return {
            // eslint-disable-next-line camelcase
            conversation_id: 1
          };
        case 'https://api.mendable.ai/v1/mendableChat':
          questionsAsked++;
          if (options.data.question === 'Hello') {
            return {
              answer: {
                text: 'Hello back'
              },
              sources: []
            };
          }
          if (options.data.question === 'Follow up') {
            return {
              answer: {
                text: 'Hello again'
              },
              sources: []
            };
          }
          break;
        case 'https://api.mendable.ai/v1/endConversation':
          return {};
      }
      throw `Invalid request: ${options.url}`;
    });
    sinon.stub(prompt, 'forInput').resolves('Follow up');
    sinon.stub(prompt, 'forSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      if (config.message === 'What would you like to do next?') {
        if (questionsAsked === 1) {
          return 'ask';
        }
        else {
          return 'end';
        }
      }

      throw `Prompt not found for '${config.message}'`;
    });
    await assert.doesNotReject(chili.startConversation(['Hello']));
  });

  it('for a follow-up question, includes the history', async () => {
    let questionsAsked = 0;
    sinon.stub(request, 'post').callsFake(async options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v1/newConversation':
          return {
            // eslint-disable-next-line camelcase
            conversation_id: 1
          };
        case 'https://api.mendable.ai/v1/mendableChat':
          questionsAsked++;
          if (options.data.question === 'Hello') {
            return {
              answer: {
                text: 'Hello back'
              },
              sources: []
            };
          }
          if (options.data.question === 'Follow up' &&
            options.data.history[0].prompt === 'Hello' &&
            options.data.history[0].response === 'Hello back') {
            return {
              answer: {
                text: 'Hello again'
              },
              sources: []
            };
          }
          break;
        case 'https://api.mendable.ai/v1/endConversation':
          return {};
      }
      throw `Invalid request: ${options.url}`;
    });
    sinon.stub(prompt, 'forInput').resolves('Follow up');
    sinon.stub(prompt, 'forSelection').callsFake(async (config) => {
      if (config.message === 'What would you like to do next?' && questionsAsked >= 2) {
        return 'end';
      }

      return 'ask';
    });
    await assert.doesNotReject(chili.startConversation(['Hello']));
  });

  it('allows ending conversation after a response', async () => {
    sinon.stub(request, 'post').callsFake(async options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v1/newConversation':
          return {
            // eslint-disable-next-line camelcase
            conversation_id: 1
          };
        case 'https://api.mendable.ai/v1/mendableChat':
          if (options.data.question === 'Hello') {
            return {
              answer: {
                text: 'Hello back'
              },
              sources: []
            };
          }
          break;
        case 'https://api.mendable.ai/v1/endConversation':
          return {};
      }
      throw `Invalid request: ${options.url}`;
    });
    sinon.stub(prompt, 'forInput').resolves('Hello world');
    sinon.stub(prompt, 'forSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      if (config.message === 'What would you like to do next?') {
        return 'end';
      }

      throw `Prompt not found for '${config.message}'`;
    });
    await assert.doesNotReject(chili.startConversation(['Hello']));
  });

  it('allows starting a new conversation after a response', async () => {
    let conversationsStarted = 0;
    sinon.stub(request, 'post').callsFake(async options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v1/newConversation':
          return {
            // eslint-disable-next-line camelcase
            conversation_id: ++conversationsStarted
          };
        case 'https://api.mendable.ai/v1/mendableChat':
          if (options.data.question === 'Hello') {
            return {
              answer: {
                text: 'Hello back'
              },
              sources: []
            };
          }
          if (options.data.question === 'Hello 2' &&
            options.data.conversation_id === 2) {
            return {
              answer: {
                text: 'Hello there'
              },
              sources: []
            };
          }
          break;
        case 'https://api.mendable.ai/v1/endConversation':
          return {};
      }
      throw `Invalid request: ${options.url}`;
    });
    sinon.stub(prompt, 'forInput').resolves('Hello 2');
    sinon.stub(prompt, 'forSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      if (config.message === 'What would you like to do next?' && conversationsStarted === 1) {
        return 'new';
      }
      if (config.message === 'What would you like to do next?' && conversationsStarted > 1) {
        return 'end';
      }

      throw `Prompt not found for '${config.message}'`;
    });
    await assert.doesNotReject(chili.startConversation(['Hello']));
  });

  it('throws exception when getting conversation ID failed', async () => {
    sinon.stub(request, 'post').callsFake(async options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v1/newConversation':
          return Promise.reject('An error has occurred');
      }
      throw `Invalid request: ${options.url}`;
    });
    sinon.stub(prompt, 'forSelection').resolves({});
    assert.rejects(chili.startConversation(['Hello']), 'An error has occurred');
  });

  it('throw exception when calling Mendable API to search in CLI docs failed', () => {
    sinon.stub(request, 'post').callsFake(async options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v1/newConversation':
          return {
            // eslint-disable-next-line camelcase
            conversation_id: 1
          };
        case 'https://api.mendable.ai/v1/mendableChat':
          throw 'An error has occurred';
      }
      throw `Invalid request: ${options.url}`;
    });
    sinon.stub(prompt, 'forSelection').resolves({});
    assert.rejects(chili.startConversation(['Hello']), 'An error has occurred');
  });

  it('shows help when requested using --help', async () => {
    await chili.startConversation(['--help']);
    assert(consoleLogSpy.getCalls().some(call => call.args[0].includes('CLI ASSISTANT (CHILI)')));
  });

  it('shows help when requested using -h', async () => {
    await chili.startConversation(['-h']);
    assert(consoleLogSpy.getCalls().some(call => call.args[0].includes('CLI ASSISTANT (CHILI)')));
  });

  it(`when requested help, doesn't start conversation`, async () => {
    const requestSpy = sinon.spy(request, 'post');
    await chili.startConversation(['--help']);
    assert(requestSpy.notCalled);
  });

  it('shows error message when help file not found', async () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    await chili.startConversation(['--help']);
    assert(consoleErrorSpy.calledWith('Help file not found'));
  });
});