import * as assert from 'assert';
import { sinonUtil } from '../utils/sinonUtil';

describe('utils/sinonUtil', () => {
  it('doesn\'t fail when restoring stub if the passed object is undefined', () => {
    sinonUtil.restore(undefined);
    assert(true);
  });
});