import { Log } from '@microsoft/sp-core-library';
import { override } from '@microsoft/decorators';
import * as React from 'react';

import styles from './HelloWorld.module.scss';

export interface IHelloWorldProps {
  text: string;
}

const LOG_SOURCE: string = 'HelloWorld';

export default class HelloWorld extends React.Component<IHelloWorldProps, {}> {
  @override
  public componentDidMount(): void {
    Log.info(LOG_SOURCE, 'React Element: HelloWorld mounted');
  }

  @override
  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, 'React Element: HelloWorld unmounted');
  }

  @override
  public render(): React.ReactElement<{}> {
    return (
      <div className={styles.cell}>
        { this.props.text }
      </div>
    );
  }
}
