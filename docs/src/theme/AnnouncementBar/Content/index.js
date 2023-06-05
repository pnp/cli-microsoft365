"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const react_1 = require("react");
const clsx_1 = require("clsx");
const theme_common_1 = require("@docusaurus/theme-common");
const styles_module_css_1 = require("./styles.module.css");
function AnnouncementBarContent(props) {
    const { announcementBar } = (0, theme_common_1.useThemeConfig)();
    const { content } = announcementBar;
    return (<div {...props} className={(0, clsx_1.default)(styles_module_css_1.default.content, props.className)} 
    // Developer provided the HTML, so assume it's safe.
    // eslint-disable-next-line react/no-danger
    dangerouslySetInnerHTML={{ __html: content }}/>);
}
exports.default = AnnouncementBarContent;
//# sourceMappingURL=index.js.map