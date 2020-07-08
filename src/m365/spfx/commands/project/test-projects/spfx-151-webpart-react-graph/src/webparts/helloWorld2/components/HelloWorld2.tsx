import * as React from 'react';
import styles from './HelloWorld2.module.scss';
import { IHelloWorld2Props } from './IHelloWorld2Props';
import { escape } from '@microsoft/sp-lodash-subset';

export default class HelloWorld2 extends React.Component<IHelloWorld2Props, {}> {
  public render(): React.ReactElement<IHelloWorld2Props> {
    this.props.graphClient
      .api('me')
      .get()
      .then((res: any) => {

      });

    return (
      <div className={ styles.helloWorld2 }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
