import { FN021002_PKG_spfx_deps_use_exact_version } from "./rules/FN021002_PKG_spfx_deps_use_exact_version.js";
import { FN021003_PKG_spfx_deps_installed_as_deps } from "./rules/FN021003_PKG_spfx_deps_installed_as_deps.js";
import { FN021004_PKG_spfx_devdeps_installed_as_devdeps } from "./rules/FN021004_PKG_spfx_devdeps_installed_as_devdeps.js";
import { FN021005_PKG_types_installed_as_devdep } from "./rules/FN021005_PKG_types_installed_as_devdep.js";
import { FN021006_PKG_rush_stack_compiler_installed_as_devdep } from "./rules/FN021006_PKG_rush_stack_compiler_installed_as_devdep.js";
import { FN021007_PKG_only_one_rush_stack_compiler_installed } from "./rules/FN021007_PKG_only_one_rush_stack_compiler_installed.js";
import { FN021008_PKG_no_duplicate_deps } from "./rules/FN021008_PKG_no_duplicate_deps.js";
import { FN021009_PKG_no_duplicate_oui_deps } from "./rules/FN021009_PKG_no_duplicate_oui_deps.js";
import { FN021010_PKG_gulp_installed_as_devdep } from "./rules/FN021010_PKG_gulp_installed_as_devdep.js";
import { FN021011_PKG_ajv_installed_as_devdep } from "./rules/FN021011_PKG_ajv_installed_as_devdep.js";
import { FN021012_PKG_no_duplicate_pnpjs_deps } from "./rules/FN021012_PKG_no_duplicate_pnpjs_deps.js";

export const rules = [
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