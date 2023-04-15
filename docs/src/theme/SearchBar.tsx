import React from "react";
import { MendableFloatingButton } from "@mendable/search";
import useDocusaurusContext from "@docusaurus/useDocusaurusContext";

export default function SearchBarWrapper(): JSX.Element {
  const {
    siteConfig: { customFields }
  } = useDocusaurusContext();

  const style = { darkMode: false, accentColor: "#ef5552" };

  const floatingButtonStyle = {
    color: "#fff",
    backgroundColor: "#ef5552"
  };

  return (
    <div className="mendable-search">
      <MendableFloatingButton
        anon_key={customFields.mendableAnonKey}
        style={style} 
        floatingButtonStyle={floatingButtonStyle} />
    </div>
  );
}