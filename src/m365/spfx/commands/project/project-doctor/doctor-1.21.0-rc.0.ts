import { FN002021_DEVDEP_rushstack_eslint_config } from '../project-upgrade/rules/FN002021_DEVDEP_rushstack_eslint_config.js';
import { FN001008_DEP_react } from './rules/FN001008_DEP_react.js';
import { FN001009_DEP_react_dom } from './rules/FN001009_DEP_react_dom.js';
import { FN001035_DEP_fluentui_react } from './rules/FN001035_DEP_fluentui_react.js';
import { FN002004_DEVDEP_gulp } from './rules/FN002004_DEVDEP_gulp.js';
import { FN002007_DEVDEP_ajv } from './rules/FN002007_DEVDEP_ajv.js';
import { FN002013_DEVDEP_types_webpack_env } from './rules/FN002013_DEVDEP_types_webpack_env.js';
import { FN002015_DEVDEP_types_react } from './rules/FN002015_DEVDEP_types_react.js';
import { FN002016_DEVDEP_types_react_dom } from './rules/FN002016_DEVDEP_types_react_dom.js';
import { FN002019_DEVDEP_microsoft_rush_stack_compiler } from './rules/FN002019_DEVDEP_microsoft_rush_stack_compiler.js';
import { FN021001_PKG_spfx_deps_versions_match_project_version } from './rules/FN021001_PKG_spfx_deps_versions_match_project_version.js';

export default [
  new FN001008_DEP_react('17'),
  new FN001009_DEP_react_dom('17'),
  new FN001035_DEP_fluentui_react('^8.106.4'),
  new FN002004_DEVDEP_gulp('4.0.2'),
  new FN002007_DEVDEP_ajv('^6.12.5'),
  new FN002013_DEVDEP_types_webpack_env('~1.15.2'),
  new FN002015_DEVDEP_types_react('17'),
  new FN002016_DEVDEP_types_react_dom('17'),
  new FN002019_DEVDEP_microsoft_rush_stack_compiler(['5.3']),
  new FN002021_DEVDEP_rushstack_eslint_config('4.0.1'),
  new FN021001_PKG_spfx_deps_versions_match_project_version(true)
];