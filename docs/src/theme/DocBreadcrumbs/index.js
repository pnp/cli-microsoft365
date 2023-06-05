"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const react_1 = require("react");
const clsx_1 = require("clsx");
const theme_common_1 = require("@docusaurus/theme-common");
const internal_1 = require("@docusaurus/theme-common/internal");
const Link_1 = require("@docusaurus/Link");
const Translate_1 = require("@docusaurus/Translate");
const Home_1 = require("@theme/DocBreadcrumbs/Items/Home");
const styles_module_css_1 = require("./styles.module.css");
// TODO move to design system folder
function BreadcrumbsItemLink({ children, href, isLast }) {
    const className = 'breadcrumbs__link';
    if (isLast) {
        return (<span className={className} itemProp="name">
        {children}
      </span>);
    }
    return href ? (<Link_1.default className={className} href={href} itemProp="item">
      <span itemProp="name">{children}</span>
    </Link_1.default>) : (
    // TODO Google search console doesn't like breadcrumb items without href.
    // The schema doesn't seem to require `id` for each `item`, although Google
    // insist to infer one, even if it's invalid. Removing `itemProp="item
    // name"` for now, since I don't know how to properly fix it.
    // See https://github.com/facebook/docusaurus/issues/7241
    <span className={className}>{children}</span>);
}
// TODO move to design system folder
function BreadcrumbsItem({ children, active, index, addMicrodata }) {
    return (<li {...(addMicrodata && {
        itemScope: true,
        itemProp: 'itemListElement',
        itemType: 'https://schema.org/ListItem'
    })} className={(0, clsx_1.default)('breadcrumbs__item', {
            'breadcrumbs__item--active': active
        })}>
      {children}
      <meta itemProp="position" content={String(index + 1)}/>
    </li>);
}
//! Custom function
function getCurrentHeadFolderName(folderId) {
    switch (folderId) {
        case 'about':
            return 'About';
        case 'cmd':
            return 'Commands';
        case 'concepts':
            return 'Concepts';
        case 'sample-scripts':
            return 'Sample Scripts';
        case 'user-guide':
            return 'User Guide';
        default:
            return '';
    }
}
function DocBreadcrumbs() {
    var _a, _b;
    const breadcrumbs = (0, internal_1.useSidebarBreadcrumbs)();
    const homePageRoute = (0, internal_1.useHomePageRoute)();
    if (!breadcrumbs) {
        return null;
    }
    //! Custom const
    const headFolderId = (_b = (_a = breadcrumbs[breadcrumbs.length - 1]) === null || _a === void 0 ? void 0 : _a.docId) === null || _b === void 0 ? void 0 : _b.split('/')[0];
    return (<nav className={(0, clsx_1.default)(theme_common_1.ThemeClassNames.docs.docBreadcrumbs, styles_module_css_1.default.breadcrumbsContainer)} aria-label={(0, Translate_1.translate)({
            id: 'theme.docs.breadcrumbs.navAriaLabel',
            message: 'Breadcrumbs',
            description: 'The ARIA label for the breadcrumbs',
        })}>
      <ul className="breadcrumbs" itemScope itemType="https://schema.org/BreadcrumbList">
        {homePageRoute && <Home_1.default />}
        {headFolderId &&
            getCurrentHeadFolderName(headFolderId) !== '' &&
            <BreadcrumbsItem key={999} active={false} index={999} addMicrodata={false}>
            <BreadcrumbsItemLink isLast={false}>
              {getCurrentHeadFolderName(headFolderId)}
            </BreadcrumbsItemLink>
          </BreadcrumbsItem>}
        {breadcrumbs.map((item, idx) => {
            const isLast = idx === breadcrumbs.length - 1;
            return (<BreadcrumbsItem key={idx} active={isLast} index={idx} addMicrodata={!!item.href}>
              <BreadcrumbsItemLink href={item.href} isLast={isLast}>
                {item.label}
              </BreadcrumbsItemLink>
            </BreadcrumbsItem>);
        })}
      </ul>
    </nav>);
}
exports.default = DocBreadcrumbs;
//# sourceMappingURL=index.js.map