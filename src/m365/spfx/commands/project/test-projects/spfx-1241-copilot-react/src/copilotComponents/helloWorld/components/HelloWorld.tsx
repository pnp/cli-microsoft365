import * as React from 'react';
import {
  FluentProvider,
  IdPrefixProvider,
  webLightTheme,
  webDarkTheme,
  Title3,
  Body1,
  Badge,
  Button,
  Card,
  CardHeader,
  makeStyles,
  tokens
} from '@fluentui/react-components';
import {
  ArrowExpand24Regular,
  Open24Regular,
  Chat24Regular,
  ResizeLarge24Regular
} from '@fluentui/react-icons';
import { createCopilotTextContent } from '@microsoft/sp-copilot-component';

import type { IHelloWorldProps } from './IHelloWorldProps';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
    padding: tokens.spacingHorizontalM
  },
  header: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS
  },
  badges: {
    display: 'flex',
    gap: tokens.spacingHorizontalS,
    flexWrap: 'wrap'
  },
  actions: {
    display: 'flex',
    gap: tokens.spacingHorizontalS,
    flexWrap: 'wrap'
  }
});

const EXPANDED_WIDTH: number = 600;
const EXPANDED_HEIGHT: number = 400;
const COMPACT_WIDTH: number = 400;
const COMPACT_HEIGHT: number = 250;

/**
 * Main React UI for the Copilot Component starter template.
 *
 * Demonstrates:
 * - **Theming** — wraps content in `<FluentProvider>` with a theme derived
 *   from the host's `hostContext.theme` (`'light' | 'dark'`).
 * - **Host context** — surfaces the current display mode, theme, and
 *   available display modes as live badges.
 * - **Bridge actions** — four buttons that exercise different bridge methods
 *   to show how a component communicates with the Copilot host.
 */
export default function HelloWorld(props: IHelloWorldProps): React.ReactElement {
  const {
    message, userDisplayName, siteTitle, siteUrl, hostContext,
    bridge, onRequestDisplayMode, onRequestSizeChange, strings
  } = props;
  const styles = useStyles();

  const [isExpanded, setIsExpanded] = React.useState<boolean>(false);

  const theme = hostContext.theme === 'dark' ? webDarkTheme : webLightTheme;

  // Request the Copilot host to switch this component to fullscreen mode.
  const handleExpand = React.useCallback(async (): Promise<void> => {
    await onRequestDisplayMode('fullscreen');
  }, [onRequestDisplayMode]);

  // Ask the host to open the site URL in the user's browser.
  const handleOpenLink = React.useCallback(async (): Promise<void> => {
    await bridge.openLinkAsync(siteUrl);
  }, [bridge, siteUrl]);

  // Send a follow-up message into the Copilot conversation on behalf of the user.
  // This triggers Copilot to respond, demonstrating how a component can drive the chat.
  const handleFollowUp = React.useCallback(async (): Promise<void> => {
    await bridge.sendFollowUpMessageAsync([
      createCopilotTextContent(strings.FollowUpMessage.replace('{0}', siteTitle))
    ]);
  }, [bridge, siteTitle, strings]);

  // Toggle between compact and expanded sizes by requesting a resize from the host.
  const handleResize = React.useCallback(async (): Promise<void> => {
    if (isExpanded) {
      await onRequestSizeChange(COMPACT_WIDTH, COMPACT_HEIGHT);
    } else {
      await onRequestSizeChange(EXPANDED_WIDTH, EXPANDED_HEIGHT);
    }
    setIsExpanded(!isExpanded);
  }, [onRequestSizeChange, isExpanded]);

  return (
    <IdPrefixProvider value="copilot-component-">
      <FluentProvider theme={theme} targetDocument={props.targetDocument} style={{ minHeight: '100%' }}>
        <div className={styles.root}>
          <Card>
            <CardHeader
              header={<Title3>{strings.GreetingPrefix} {userDisplayName}!</Title3>}
              description={<Body1>{message}</Body1>}
            />

            <div className={styles.badges}>
              <Badge appearance="outline">{strings.SiteBadgePrefix} {siteTitle}</Badge>
              <Badge appearance="outline" color="informative">
                {strings.ThemeBadgePrefix} {hostContext.theme || strings.UnknownTheme}
              </Badge>
              <Badge appearance="outline" color="subtle">
                {strings.ModeBadgePrefix} {hostContext.displayMode || strings.DefaultDisplayMode}
              </Badge>
            </div>
          </Card>

          <div className={styles.actions}>
            <Button appearance="primary" icon={<ArrowExpand24Regular />} onClick={handleExpand}>
              {strings.ExpandButtonLabel}
            </Button>
            <Button appearance="secondary" icon={<Open24Regular />} onClick={handleOpenLink}>
              {strings.OpenSiteButtonLabel}
            </Button>
            <Button appearance="secondary" icon={<Chat24Regular />} onClick={handleFollowUp}>
              {strings.FollowUpButtonLabel}
            </Button>
            <Button appearance="secondary" icon={<ResizeLarge24Regular />} onClick={handleResize}>
              {isExpanded ? strings.CompactButtonLabel : strings.ResizeButtonLabel}
            </Button>
          </div>
        </div>
      </FluentProvider>
    </IdPrefixProvider>
  );
}
