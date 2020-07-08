import * as ko from 'knockout';
import styles from './HelloWorld.module.scss';
import { IHelloWorldWebPartProps } from './HelloWorldWebPart';

export interface IHelloWorldBindingContext extends IHelloWorldWebPartProps {
  shouter: KnockoutSubscribable<{}>;
}

export default class HelloWorldViewModel {
  public description: KnockoutObservable<string> = ko.observable('');

  public helloWorldClass: string = styles.helloWorld;
  public containerClass: string = styles.container;
  public rowClass: string = styles.row;
  public columnClass: string = styles.column;
  public titleClass: string = styles.title;
  public subTitleClass: string = styles.subTitle;
  public descriptionClass: string = styles.description;
  public buttonClass: string = styles.button;
  public labelClass: string = styles.label;

  constructor(bindings: IHelloWorldBindingContext) {
    this.description(bindings.description);

    // When web part description is updated, change this view model's description.
    bindings.shouter.subscribe((value: string) => {
      this.description(value);
    }, this, 'description');
  }
}
