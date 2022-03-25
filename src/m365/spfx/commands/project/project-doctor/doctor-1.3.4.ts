import { FN001008_DEP_react } from './rules/FN001008_DEP_react';
import { FN001009_DEP_react_dom } from './rules/FN001009_DEP_react_dom';
import { FN002004_DEVDEP_gulp } from './rules/FN002004_DEVDEP_gulp';
import { FN002007_DEVDEP_ajv } from './rules/FN002007_DEVDEP_ajv';
import { FN002013_DEVDEP_types_webpack_env } from './rules/FN002013_DEVDEP_types_webpack_env';
import { FN002015_DEVDEP_types_react } from './rules/FN002015_DEVDEP_types_react';
import { FN002016_DEVDEP_types_react_dom } from './rules/FN002016_DEVDEP_types_react_dom';

module.exports = [
  new FN001008_DEP_react('15'),
  new FN001009_DEP_react_dom('15'),
  new FN002004_DEVDEP_gulp('~3.9.1'),
  new FN002007_DEVDEP_ajv('~5.2.2'),
  new FN002013_DEVDEP_types_webpack_env('>=1.12.1 <1.14.0'),
  new FN002015_DEVDEP_types_react('15'),
  new FN002016_DEVDEP_types_react_dom('15')
];