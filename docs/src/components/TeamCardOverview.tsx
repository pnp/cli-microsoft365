import React from 'react';
import useBaseUrl from '@docusaurus/useBaseUrl';
import styles from '../scss/TeamCard.module.scss';

interface IIndividual {
  name: string;
  company?: string;
  github?: string;
  twitter?: string;
}

interface ITeamCardOverview {
  individuals: IIndividual[];
}

const TeamCardOverview = ({individuals}: ITeamCardOverview): JSX.Element => (
  <div className={styles.grid}>
    {
      individuals.map(individual => 
        <div className={styles.gridItemContainer}>
          <div className={styles.gridItem}>
            <div className={styles.gridItemAlignCenter}>
              <img alt='GitHub avatar' src={individual.github ? `https://github.com/${individual.github}.png` : `https://ui-avatars.com/api/?name=${individual.name}`} className={styles.gridItemImg} />
            </div>
            <div className={styles.gridItemAlignCenter}>
              <div className={styles.gridItemText}>
                <div className={styles.gridItemName}>{individual.name}</div>
                <div className={styles.gridItemCompany}>{individual.company}</div>
              </div>
            </div>
            <div className={styles.gridItemAlignCenter}>
              {
                individual.github
                &&
                <a href={`https://github.com/${individual.github}`} title='GitHub' className={styles.gridItemLink}>
                  <img alt='GitHub' src={useBaseUrl('/img/github-icon.svg')} className={styles.gridItemLinkImg} />
                </a>
              }
              {
                individual.twitter
                &&
                <a href={`https://twitter.com/${individual.twitter}`} title='Twitter' className={styles.gridItemLink}>
                  <img alt='Twitter' src={useBaseUrl('/img/twitter-icon.svg')} className={styles.gridItemLinkImg} />
                </a>
              }
            </div>
          </div>
        </div>
      )
    }
  </div>
);

export default TeamCardOverview;