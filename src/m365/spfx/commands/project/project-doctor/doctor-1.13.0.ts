import { FN001008_DEP_react } from './rules/FN001008_DEP_react.js';
import { FN001009_DEP_react_dom } from './rules/FN001009_DEP_react_dom.js';
import { FN001022_DEP_office_ui_fabric_react } from './rules/FN001022_DEP_office_ui_fabric_react.js';
import { FN002004_DEVDEP_gulp } from './rules/FN002004_DEVDEP_gulp.js';
import { FN002007_DEVDEP_ajv } from './rules/FN002007_DEVDEP_ajv.js';
import { FN002013_DEVDEP_types_webpack_env } from './rules/FN002013_DEVDEP_types_webpack_env.js';
import { FN002015_DEVDEP_types_react } from './rules/FN002015_DEVDEP_types_react.js';
import { FN002016_DEVDEP_types_react_dom } from './rules/FN002016_DEVDEP_types_react_dom.js';
import { FN002019_DEVDEP_microsoft_rush_stack_compiler } from './rules/FN002019_DEVDEP_microsoft_rush_stack_compiler.js';

export default [
  new FN001008_DEP_react('16'),
  new FN001009_DEP_react_dom('16'),
  new FN001022_DEP_office_ui_fabric_react('7.174.1'),
  new FN002004_DEVDEP_gulp('~4.0.2'),
  new FN002007_DEVDEP_ajv('~5.2.2'),
  new FN002013_DEVDEP_types_webpack_env('1.13.1'),
  new FN002015_DEVDEP_types_react('16'),
  new FN002016_DEVDEP_types_react_dom('16'),
  new FN002019_DEVDEP_microsoft_rush_stack_compiler(['3.9'])
];