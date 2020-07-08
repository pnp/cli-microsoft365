import * as ko from 'knockout';
import styles from './HelloWorld.module.scss';
import { IHelloWorldWebPartProps } from './IHelloWorldWebPartProps';

export interface IHelloWorldBindingContext extends IHelloWorldWebPartProps {
  shouter: KnockoutSubscribable<{}>;
}

export default class HelloWorldViewModel {
  public description: KnockoutObservable<string> = ko.observable('');

  public labelClass: string = styles.label;
  public helloWorldClass: string = styles.helloWorld;
  public containerClass: string = styles.container;
  public rowClass: string = `ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}`;
  public buttonClass: string = `ms-Button ${styles.button}`;

  constructor(bindings: IHelloWorldBindingContext) {
    this.description(bindings.description);

    // When web part description is updated, change this view model's description.
    bindings.shouter.subscribe((value: string) => {
      this.description(value);
    }, this, 'description');
  }
}
