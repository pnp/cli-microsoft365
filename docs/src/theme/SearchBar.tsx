import React from 'react';
import { MendableFloatingButton } from '@mendable/search';
import { useColorMode } from '@docusaurus/theme-common';
import useDocusaurusContext from '@docusaurus/useDocusaurusContext';

export default function SearchBarWrapper(): JSX.Element {
  const {
    siteConfig: { customFields }
  } = useDocusaurusContext();

  return (
    <MendableFloatingButton
      anon_key={customFields.mendableAnonKey as string}
      style={{
        darkMode: useColorMode().isDarkTheme,
        accentColor: '#ef5552'
      }} 
      floatingButtonStyle={{
        color: '#fff',
        backgroundColor: '#ef5552'
      }}
    />
  );
}