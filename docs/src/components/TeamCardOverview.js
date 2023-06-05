"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const react_1 = require("react");
const useBaseUrl_1 = require("@docusaurus/useBaseUrl");
const TeamCard_module_scss_1 = require("../scss/TeamCard.module.scss");
const TeamCardOverview = ({ individuals }) => (<div className={TeamCard_module_scss_1.default.grid}>
    {individuals.map(individual => <div className={TeamCard_module_scss_1.default.gridItemContainer}>
          <div className={TeamCard_module_scss_1.default.gridItem}>
            <div className={TeamCard_module_scss_1.default.gridItemAlignCenter}>
              <img alt='GitHub avatar' src={individual.github ? `https://github.com/${individual.github}.png` : `https://ui-avatars.com/api/?name=${individual.name}`} className={TeamCard_module_scss_1.default.gridItemImg}/>
            </div>
            <div className={TeamCard_module_scss_1.default.gridItemAlignCenter}>
              <div className={TeamCard_module_scss_1.default.gridItemText}>
                <div className={TeamCard_module_scss_1.default.gridItemName}>{individual.name}</div>
                <div className={TeamCard_module_scss_1.default.gridItemCompany}>{individual.company}</div>
              </div>
            </div>
            <div className={TeamCard_module_scss_1.default.gridItemAlignCenter}>
              {individual.github
            &&
                <a href={`https://github.com/${individual.github}`} title='GitHub' className={TeamCard_module_scss_1.default.gridItemLink}>
                  <img alt='GitHub' src={(0, useBaseUrl_1.default)('/img/github-icon.svg')} className={TeamCard_module_scss_1.default.gridItemLinkImg}/>
                </a>}
              {individual.twitter
            &&
                <a href={`https://twitter.com/${individual.twitter}`} title='Twitter' className={TeamCard_module_scss_1.default.gridItemLink}>
                  <img alt='Twitter' src={(0, useBaseUrl_1.default)('/img/twitter-icon.svg')} className={TeamCard_module_scss_1.default.gridItemLinkImg}/>
                </a>}
            </div>
          </div>
        </div>)}
  </div>);
exports.default = TeamCardOverview;
//# sourceMappingURL=TeamCardOverview.js.map