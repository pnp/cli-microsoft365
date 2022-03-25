import { FN021001_PKG_spfx_deps_versions_match_project_version } from "./rules/FN021001_PKG_spfx_deps_versions_match_project_version";
import { FN021002_PKG_spfx_deps_use_exact_version } from "./rules/FN021002_PKG_spfx_deps_use_exact_version";
import { FN021003_PKG_spfx_deps_installed_as_deps } from "./rules/FN021003_PKG_spfx_deps_installed_as_deps";
import { FN021004_PKG_spfx_devdeps_installed_as_devdeps } from "./rules/FN021004_PKG_spfx_devdeps_installed_as_devdeps";
import { FN021005_PKG_types_installed_as_devdep } from "./rules/FN021005_PKG_types_installed_as_devdep";
import { FN021006_PKG_rush_stack_compiler_installed_as_devdep } from "./rules/FN021006_PKG_rush_stack_compiler_installed_as_devdep";
import { FN021007_PKG_only_one_rush_stack_compiler_installed } from "./rules/FN021007_PKG_only_one_rush_stack_compiler_installed";
import { FN021008_PKG_no_duplicate_deps } from "./rules/FN021008_PKG_no_duplicate_deps";
import { FN021009_PKG_no_duplicate_oui_deps } from "./rules/FN021009_PKG_no_duplicate_oui_deps";
import { FN021010_PKG_gulp_installed_as_devdep } from "./rules/FN021010_PKG_gulp_installed_as_devdep";
import { FN021011_PKG_ajv_installed_as_devdep } from "./rules/FN021011_PKG_ajv_installed_as_devdep";
import { FN021012_PKG_no_duplicate_pnpjs_deps } from "./rules/FN021012_PKG_no_duplicate_pnpjs_deps";

export const rules = [
  new FN021001_PKG_spfx_deps_versions_match_project_version(),
  new FN021002_PKG_spfx_deps_use_exact_version(),
  new FN021003_PKG_spfx_deps_installed_as_deps(),
  new FN021004_PKG_spfx_devdeps_installed_as_devdeps(),
  new FN021005_PKG_types_installed_as_devdep(),
  new FN021006_PKG_rush_stack_compiler_installed_as_devdep(),
  new FN021007_PKG_only_one_rush_stack_compiler_installed(),
  new FN021008_PKG_no_duplicate_deps(),
  new FN021009_PKG_no_duplicate_oui_deps(),
  new FN021010_PKG_gulp_installed_as_devdep(),
  new FN021011_PKG_ajv_installed_as_devdep(),
  new FN021012_PKG_no_duplicate_pnpjs_deps()
];