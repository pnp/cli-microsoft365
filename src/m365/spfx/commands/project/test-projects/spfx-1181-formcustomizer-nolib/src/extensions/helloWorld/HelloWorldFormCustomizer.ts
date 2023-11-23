import { Log } from '@microsoft/sp-core-library';
import {
  BaseFormCustomizer
} from '@microsoft/sp-listview-extensibility';

import styles from './HelloWorldFormCustomizer.module.scss';

/**
 * If your form customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IHelloWorldFormCustomizerProperties {
  // This is an example; replace with your own property
  sampleText?: string;
}

const LOG_SOURCE: string = 'HelloWorldFormCustomizer';

export default class HelloWorldFormCustomizer
  extends BaseFormCustomizer<IHelloWorldFormCustomizerProperties> {

  public onInit(): Promise<void> {
    // Add your custom initialization to this method. The framework will wait
    // for the returned promise to resolve before rendering the form.
    Log.info(LOG_SOURCE, 'Activated HelloWorldFormCustomizer with properties:');
    Log.info(LOG_SOURCE, JSON.stringify(this.properties, undefined, 2));
    return Promise.resolve();
  }

  public render(): void {
    // Use this method to perform your custom rendering.
    this.domElement.innerHTML = `<div class="${ styles.helloWorld }"></div>`;
  }

  public onDispose(): void {
    // This method should be used to free any resources that were allocated during rendering.
    super.onDispose();
  }

  //@ts-ignore
  private _onSave = (): void => {
    // TODO: Add your custom save logic here.

    // You MUST call this.formSaved() after you save the form.
    this.formSaved();
  }

  //@ts-ignore
  private _onClose = (): void => {
    // TODO: Add your custom close logic here.

    // You MUST call this.formClosed() after you close the form.
    this.formClosed();
  }
}
