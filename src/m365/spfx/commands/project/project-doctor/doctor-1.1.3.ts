import { FN001008_DEP_react } from './rules/FN001008_DEP_react.js';
import { FN001009_DEP_react_dom } from './rules/FN001009_DEP_react_dom.js';
import { FN002004_DEVDEP_gulp } from './rules/FN002004_DEVDEP_gulp.js';
import { FN002013_DEVDEP_types_webpack_env } from './rules/FN002013_DEVDEP_types_webpack_env.js';
import { FN002015_DEVDEP_types_react } from './rules/FN002015_DEVDEP_types_react.js';
import { FN002016_DEVDEP_types_react_dom } from './rules/FN002016_DEVDEP_types_react_dom.js';

export default [
  new FN001008_DEP_react('15'),
  new FN001009_DEP_react_dom('15'),
  new FN002004_DEVDEP_gulp('~3.9.1'),
  new FN002013_DEVDEP_types_webpack_env('>=1.12.1 <1.14.0'),
  new FN002015_DEVDEP_types_react('0.14'),
  new FN002016_DEVDEP_types_react_dom('0.14')
];