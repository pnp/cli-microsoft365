"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const react_1 = require("react");
const theme_common_1 = require("@docusaurus/theme-common");
const internal_1 = require("@docusaurus/theme-common/internal");
const CloseButton_1 = require("@theme/AnnouncementBar/CloseButton");
const Content_1 = require("@theme/AnnouncementBar/Content");
const styles_module_css_1 = require("./styles.module.css");
function AnnouncementBar() {
    const { announcementBar } = (0, theme_common_1.useThemeConfig)();
    const { isActive, close } = (0, internal_1.useAnnouncementBar)();
    if (!isActive) {
        return null;
    }
    const { backgroundColor, textColor, isCloseable } = announcementBar;
    return (<div className={styles_module_css_1.default.announcementBar} style={{ backgroundColor, color: textColor }} role="banner">
      {isCloseable && <div className={styles_module_css_1.default.announcementBarPlaceholder}/>}
      <Content_1.default className={styles_module_css_1.default.announcementBarContent}/>
      {isCloseable && (<CloseButton_1.default onClick={close} className={styles_module_css_1.default.announcementBarClose}/>)}
    </div>);
}
exports.default = AnnouncementBar;
//# sourceMappingURL=index.js.map