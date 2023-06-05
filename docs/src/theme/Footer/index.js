"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const react_1 = require("react");
const clsx_1 = require("clsx");
const styles_module_scss_1 = require("./styles.module.scss");
const github_icon_svg_1 = require("@site/static/img/github-icon.svg");
const discord_icon_svg_1 = require("@site/static/img/discord-icon.svg");
const twitter_icon_svg_1 = require("@site/static/img/twitter-icon.svg");
const youtube_icon_svg_1 = require("@site/static/img/youtube-icon.svg");
const link_icon_svg_1 = require("@site/static/img/link-icon.svg");
const search_1 = require("@mendable/search");
const theme_common_1 = require("@docusaurus/theme-common");
const useDocusaurusContext_1 = require("@docusaurus/useDocusaurusContext");
function Footer() {
    const { siteConfig: { customFields } } = (0, useDocusaurusContext_1.default)();
    return (<div>
      <footer className={(0, clsx_1.default)('footer', 'footer--dark')}>
        <div className="container container-fluid">
          <div className={styles_module_scss_1.default.footerContainer}>
            <p className={styles_module_scss_1.default.footerPromo}>
              Built with <a href='https://docusaurus.io/' target='_blank' rel="noopener" title='Docusaurus'>Docusaurus</a>
            </p>

            <div className={styles_module_scss_1.default.footerIcons}>    
              <a href="https://github.com/pnp/cli-microsoft365" target="_blank" rel="noopener" title="GitHub">
                <github_icon_svg_1.default className={styles_module_scss_1.default.svg}/>
              </a>

              <a href="https://aka.ms/cli-m365/discord" target="_blank" rel="noopener" title="Discord">
                <discord_icon_svg_1.default className={styles_module_scss_1.default.svg}/>
              </a>

              <a href="https://twitter.com/climicrosoft365" target="_blank" rel="noopener" title="Twitter">
                <twitter_icon_svg_1.default className={styles_module_scss_1.default.svg}/>
              </a>

              <a href="http://aka.ms/sppnp-videos" target="_blank" rel="noopener" title="YouTube">
                <youtube_icon_svg_1.default className={styles_module_scss_1.default.svg}/>
              </a>

              <a href="https://aka.ms/sppnp" target="_blank" rel="noopener" title="Microsoft 365 & Power Platform Community Website">
                <link_icon_svg_1.default className={styles_module_scss_1.default.svg}/>
              </a>
            </div>
          </div>
        </div>
      </footer>

      <search_1.MendableFloatingButton anon_key={customFields.mendableAnonKey} style={{
            darkMode: (0, theme_common_1.useColorMode)().isDarkTheme,
            accentColor: '#ef5552'
        }} floatingButtonStyle={{
            color: '#fff',
            backgroundColor: '#ef5552'
        }} cmdShortcutKey='m'/>
    </div>);
}
exports.default = react_1.default.memo(Footer);
//# sourceMappingURL=index.js.map