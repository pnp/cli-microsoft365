"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const react_1 = require("react");
const clsx_1 = require("clsx");
const Translate_1 = require("@docusaurus/Translate");
const Close_1 = require("@theme/Icon/Close");
const styles_module_css_1 = require("./styles.module.css");
function AnnouncementBarCloseButton(props) {
    return (<button type="button" aria-label={(0, Translate_1.translate)({
            id: 'theme.AnnouncementBar.closeButtonAriaLabel',
            message: 'Close',
            description: 'The ARIA label for close button of announcement bar',
        })} {...props} className={(0, clsx_1.default)('clean-btn close', styles_module_css_1.default.closeButton, props.className)}>
      <Close_1.default width={14} height={14} strokeWidth={3.1}/>
    </button>);
}
exports.default = AnnouncementBarCloseButton;
//# sourceMappingURL=index.js.map