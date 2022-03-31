import { FN001008_DEP_react } from './rules/FN001008_DEP_react';
import { FN001009_DEP_react_dom } from './rules/FN001009_DEP_react_dom';
import { FN002004_DEVDEP_gulp } from './rules/FN002004_DEVDEP_gulp';
import { FN002007_DEVDEP_ajv } from './rules/FN002007_DEVDEP_ajv';
import { FN002013_DEVDEP_types_webpack_env } from './rules/FN002013_DEVDEP_types_webpack_env';
import { FN002015_DEVDEP_types_react } from './rules/FN002015_DEVDEP_types_react';
import { FN002016_DEVDEP_types_react_dom } from './rules/FN002016_DEVDEP_types_react_dom';
import { FN002019_DEVDEP_microsoft_rush_stack_compiler } from './rules/FN002019_DEVDEP_microsoft_rush_stack_compiler';

module.exports = [
  new FN001008_DEP_react('16'),
  new FN001009_DEP_react_dom('16'),
  new FN002004_DEVDEP_gulp('~3.9.1'),
  new FN002007_DEVDEP_ajv('~5.2.2'),
  new FN002013_DEVDEP_types_webpack_env('1.13.1'),
  new FN002015_DEVDEP_types_react('16'),
  new FN002016_DEVDEP_types_react_dom('16'),
  new FN002019_DEVDEP_microsoft_rush_stack_compiler(['2.7', '2.9', '3.0'])
];