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
    consoleLogSpy = sinon.stub(console, 'log').callsFake(() => { });
    consoleErrorSpy = sinon.stub(console, 'error').callsFake(() => { });
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

  it('starts a conversation using a prompt from args when specified', async () => {
    sinon.stub(request, 'post').callsFake(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: 1
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          if (options.data.question === 'Hello') {
            return Promise.resolve({
              answer: {
                text: 'Hello back'
              },
              sources: []
            });
          }
          break;
        case 'https://api.mendable.ai/v0/endConversation':
          return Promise.resolve({});
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
    await assert.doesNotReject(chili.startConversation(['Hello']));
  });

  it('starts a conversation when a prompt specified as a single arg', async () => {
    sinon.stub(request, 'post').callsFake(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: 1
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          if (options.data.question === 'Hello world') {
            return Promise.resolve({
              answer: {
                text: 'Hello back'
              },
              sources: []
            });
          }
          break;
        case 'https://api.mendable.ai/v0/endConversation':
          return Promise.resolve({});
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
    sinon.stub(request, 'post').callsFake(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: 1
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          if (options.data.question === 'Hello world') {
            return Promise.resolve({
              answer: {
                text: 'Hello back'
              },
              sources: []
            });
          }
          break;
        case 'https://api.mendable.ai/v0/endConversation':
          return Promise.resolve({});
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
    sinon.stub(request, 'post').callsFake(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: 1
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          if (options.data.question === 'Hello world') {
            return Promise.resolve({
              answer: {
                text: 'Hello back'
              },
              sources: []
            });
          }
          break;
        case 'https://api.mendable.ai/v0/endConversation':
          return Promise.resolve({});
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
    sinon.stub(request, 'post').callsFake(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: 1
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          if (options.data.question === 'Hello') {
            return Promise.resolve({
              answer: {
                text: 'Hello back'
              },
              sources: []
            });
          }
          break;
        case 'https://api.mendable.ai/v0/endConversation':
          return Promise.resolve({});
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
    sinon.stub(request, 'post').callsFake(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: 1
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          if (options.data.question === 'Hello') {
            return Promise.resolve({
              answer: {
                text: 'Hello back'
              },
              sources: []
            });
          }
          break;
        case 'https://api.mendable.ai/v0/endConversation':
          return Promise.resolve({});

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
    await chili.startConversation(['Hello', '--no-rating']);
    assert(consoleLogSpy.calledWith('Hello back'));
  });

  it('in response formats MD in terminal-friendly way', async () => {
    sinon.stub(request, 'post').callsFake(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: 1
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          if (options.data.question === 'Hello') {
            return Promise.resolve({
              answer: {
                text: 'Hello **back**'
              },
              sources: []
            });
          }
          break;
        case 'https://api.mendable.ai/v0/endConversation':
          return Promise.resolve({});
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
    sinon.stub(request, 'post').callsFake(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: 1
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          if (options.data.question === 'Hello') {
            return Promise.resolve({
              answer: {
                text: 'Hello back'
              },
              sources: [
                {
                  link: 'https://example.com/source-1'
                }
              ]
            });
          }
          break;
        case 'https://api.mendable.ai/v0/endConversation':
          return Promise.resolve({});
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

  it('prompts for rating the response when rating is enabled', async () => {
    let promptedForRating = false;
    sinon.stub(request, 'post').callsFake(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: 1
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          if (options.data.question === 'Hello world') {
            return Promise.resolve({
              answer: {
                text: 'Hello back'
              },
              // eslint-disable-next-line camelcase
              message_id: 1,
              sources: []
            });
          }
          break;
        case 'https://api.mendable.ai/v0/rateMessage':
          return Promise.resolve({});
        case 'https://api.mendable.ai/v0/endConversation':
          return Promise.resolve({});
      }
      throw `Invalid request: ${options.url}`;
    });
    sinon.stub(prompt, 'forInput').resolves('Hello world');
    sinon.stub(prompt, 'forSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      if (config.message === 'Was this helpful?') {
        promptedForRating = true;
        return 1;
      }
      else if (config.message === 'What would you like to do next?') {
        return 'end';
      }

      throw `Prompt not found for '${config.message}'`;
    });
    await chili.startConversation(['Hello world']);
    assert.strictEqual(promptedForRating, true);
  });

  it(`doesn't prompt for rating the response when rating is disabled`, async () => {
    let promptedForRating = false;
    sinon.stub(request, 'post').callsFake(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: 1
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          if (options.data.question === 'Hello world') {
            return Promise.resolve({
              answer: {
                text: 'Hello back'
              },
              // eslint-disable-next-line camelcase
              message_id: 1,
              sources: []
            });
          }
          break;
        case 'https://api.mendable.ai/v0/endConversation':
          return Promise.resolve({});
      }
      throw `Invalid request: ${options.url}`;
    });
    sinon.stub(prompt, 'forInput').resolves('Hello world');
    sinon.stub(prompt, 'forSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      if (config.message === 'Was this helpful?') {
        promptedForRating = true;
        return 1;
      }
      else if (config.message === 'What would you like to do next?') {
        return 'end';
      }

      throw `Prompt not found for '${config.message}'`;
    });
    await chili.startConversation(['Hello world', '--no-rating']);
    assert.strictEqual(promptedForRating, false);
  });

  it('sends positive rating to Mendable', async () => {
    sinon.stub(request, 'post').callsFake(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: 1
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          if (options.data.question === 'Hello world') {
            return Promise.resolve({
              answer: {
                text: 'Hello back'
              },
              // eslint-disable-next-line camelcase
              message_id: 1,
              sources: []
            });
          }
          break;
        case 'https://api.mendable.ai/v0/rateMessage':
          if (options.data.rating === 1) {
            return Promise.resolve({});
          }
          break;
        case 'https://api.mendable.ai/v0/endConversation':
          return Promise.resolve({});
      }
      throw `Invalid request: ${options.url}`;
    });
    sinon.stub(prompt, 'forInput').resolves('Hello world');
    sinon.stub(prompt, 'forSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      if (config.message === 'Was this helpful?') {
        return 1;
      }
      if (config.message === 'What would you like to do next?') {
        return 'end';
      }

      throw `Prompt not found for '${config.message}'`;
    });
    await assert.doesNotReject(chili.startConversation(['Hello world']));
  });

  it('sends negative rating to Mendable', async () => {
    sinon.stub(request, 'post').callsFake(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: 1
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          if (options.data.question === 'Hello world') {
            return Promise.resolve({
              answer: {
                text: 'Hello back'
              },
              // eslint-disable-next-line camelcase
              message_id: 1,
              sources: []
            });
          }
          break;
        case 'https://api.mendable.ai/v0/rateMessage':
          if (options.data.rating === -1) {
            return Promise.resolve({});
          }
          break;
        case 'https://api.mendable.ai/v0/endConversation':
          return Promise.resolve({});
      }
      throw `Invalid request: ${options.url}`;
    });
    sinon.stub(prompt, 'forInput').resolves('Hello world');
    sinon.stub(prompt, 'forSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      if (config.message === 'Was this helpful?') {
        return -1;
      }
      if (config.message === 'What would you like to do next?') {
        return 'end';
      }

      throw `Prompt not found for '${config.message}'`;
    });
    await assert.doesNotReject(chili.startConversation(['Hello world']));
  });

  it(`doesn't send rating to Mendable when user chose to skip`, async () => {
    sinon.stub(request, 'post').callsFake(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: 1
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          if (options.data.question === 'Hello world') {
            return Promise.resolve({
              answer: {
                text: 'Hello back'
              },
              // eslint-disable-next-line camelcase
              message_id: 1,
              sources: []
            });
          }
          break;
        case 'https://api.mendable.ai/v0/endConversation':
          return Promise.resolve({});
      }
      throw `Invalid request: ${options.url}`;
    });
    sinon.stub(prompt, 'forInput').resolves('Hello world');
    sinon.stub(prompt, 'forSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      if (config.message === 'Was this helpful?') {
        return 0;
      }
      if (config.message === 'What would you like to do next?') {
        return 'end';
      }

      throw `Prompt not found for '${config.message}'`;
    });
    await assert.doesNotReject(chili.startConversation(['Hello world']));
  });

  it(`doesn't fail when rating the response failed`, async () => {
    sinon.stub(request, 'post').callsFake(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: 1
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          if (options.data.question === 'Hello world') {
            return Promise.resolve({
              answer: {
                text: 'Hello back'
              },
              // eslint-disable-next-line camelcase
              message_id: 1,
              sources: []
            });
          }
          break;
        case 'https://api.mendable.ai/v0/rateMessage':
          return Promise.reject('An error has occurred');
        case 'https://api.mendable.ai/v0/endConversation':
          return Promise.resolve({});
      }
      throw `Invalid request: ${options.url}`;
    });
    sinon.stub(prompt, 'forInput').resolves('Hello world');
    sinon.stub(prompt, 'forSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      if (config.message === 'Was this helpful?') {
        return 1;
      }
      if (config.message === 'What would you like to do next?') {
        return 'end';
      }

      throw `Prompt not found for '${config.message}'`;
    });
    await assert.doesNotReject(chili.startConversation(['Hello world']));
  });

  it(`when rating the response failed, shows error in debug mode`, async () => {
    sinon.stub(request, 'post').callsFake(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: 1
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          if (options.data.question === 'Hello world') {
            return Promise.resolve({
              answer: {
                text: 'Hello back'
              },
              // eslint-disable-next-line camelcase
              message_id: 1,
              sources: []
            });
          }
          break;
        case 'https://api.mendable.ai/v0/rateMessage':
          return Promise.reject('An error has occurred');
        case 'https://api.mendable.ai/v0/endConversation':
          return Promise.resolve({});
      }
      throw `Invalid request: ${options.url}`;
    });
    sinon.stub(prompt, 'forInput').resolves('Hello world');
    sinon.stub(prompt, 'forSelection').callsFake(async (config: SelectionConfig<unknown>): Promise<unknown> => {
      if (config.message === 'Was this helpful?') {
        return 1;
      }
      else if (config.message === 'What would you like to do next?') {
        return 'end';
      }

      throw `Prompt not found for '${config.message}'`;
    });
    await chili.startConversation(['Hello world', '--debug']);
    assert(consoleErrorSpy.calledWith('An error has occurred while rating the response: An error has occurred'));
  });

  it('allows asking a follow-up question after a response', async () => {
    let questionsAsked = 0;
    sinon.stub(request, 'post').callsFake(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: 1
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          questionsAsked++;
          if (options.data.question === 'Hello') {
            return Promise.resolve({
              answer: {
                text: 'Hello back'
              },
              sources: []
            });
          }
          if (options.data.question === 'Follow up') {
            return Promise.resolve({
              answer: {
                text: 'Hello again'
              },
              sources: []
            });
          }
          break;
        case 'https://api.mendable.ai/v0/endConversation':
          return Promise.resolve({});
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
    await assert.doesNotReject(chili.startConversation(['Hello', '--no-rating']));
  });

  it('for a follow-up question, includes the history', async () => {
    let questionsAsked = 0;
    sinon.stub(request, 'post').callsFake(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: 1
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          questionsAsked++;
          if (options.data.question === 'Hello') {
            return Promise.resolve({
              answer: {
                text: 'Hello back'
              },
              sources: []
            });
          }
          if (options.data.question === 'Follow up' &&
            options.data.history[0].prompt === 'Hello' &&
            options.data.history[0].response === 'Hello back') {
            return Promise.resolve({
              answer: {
                text: 'Hello again'
              },
              sources: []
            });
          }
          break;
        case 'https://api.mendable.ai/v0/endConversation':
          return Promise.resolve({});
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
    await assert.doesNotReject(chili.startConversation(['Hello', '--no-rating']));
  });

  it('allows ending conversation after a response', async () => {
    sinon.stub(request, 'post').callsFake(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: 1
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          if (options.data.question === 'Hello') {
            return Promise.resolve({
              answer: {
                text: 'Hello back'
              },
              sources: []
            });
          }
          break;
        case 'https://api.mendable.ai/v0/endConversation':
          return Promise.resolve({});
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
    await assert.doesNotReject(chili.startConversation(['Hello', '--no-rating']));
  });

  it('allows starting a new conversation after a response', async () => {
    let conversationsStarted = 0;
    sinon.stub(request, 'post').callsFake(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: ++conversationsStarted
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          if (options.data.question === 'Hello') {
            return Promise.resolve({
              answer: {
                text: 'Hello back'
              },
              sources: []
            });
          }
          if (options.data.question === 'Hello 2' &&
            options.data.conversation_id === 2) {
            return Promise.resolve({
              answer: {
                text: 'Hello there'
              },
              sources: []
            });
          }
          break;
        case 'https://api.mendable.ai/v0/endConversation':
          return Promise.resolve({});
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
    await assert.doesNotReject(chili.startConversation(['Hello', '--no-rating']));
  });

  it('throws exception when getting conversation ID failed', async () => {
    sinon.stub(request, 'post').callsFake(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.reject('An error has occurred');
      }
      throw `Invalid request: ${options.url}`;
    });
    sinon.stub(prompt, 'forSelection').resolves({});
    assert.rejects(chili.startConversation(['Hello']), 'An error has occurred');
  });

  it('throw exception when calling Mendable API to search in CLI docs failed', () => {
    sinon.stub(request, 'post').callsFake(options => {
      switch (options.url) {
        case 'https://api.mendable.ai/v0/newConversation':
          return Promise.resolve({
            // eslint-disable-next-line camelcase
            conversation_id: 1
          });
        case 'https://api.mendable.ai/v0/mendableChat':
          return Promise.reject('An error has occurred');
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