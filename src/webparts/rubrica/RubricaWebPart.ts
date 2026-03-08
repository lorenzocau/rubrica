import * as React from "react";
import * as ReactDom from "react-dom";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import { getSP, getGraph } from "./pnpjsConfig";
import Rubrica from "./components/Rubrica";

// BaseClientSideWebPart è la classe base di SPFx per tutti i web part.
// Non ha props configurabili dal pannello proprietà, quindi usiamo
// un oggetto vuoto come tipo generico.
export default class RubricaWebPart extends BaseClientSideWebPart<{}> {

  // onInit viene chiamato UNA SOLA VOLTA da SPFx prima del render.
  // È l'unico posto dove this.context è disponibile e dove dobbiamo
  // inizializzare PnPjs — da questo momento getSP() e getGraph()
  // possono essere chiamati ovunque senza passare il context.
  protected async onInit(): Promise<void> {
    await super.onInit();
    getSP(this.context);
    getGraph(this.context);
  }

  // render viene chiamato da SPFx dopo onInit.
  // Il suo unico compito è montare il componente React principale
  // nel DOM element che SPFx ci mette a disposizione (this.domElement).
  // Da questo momento in poi è React a gestire tutto.
  public render(): void {
    const element = React.createElement(Rubrica);
    ReactDom.render(element, this.domElement);
  }

  // onDispose viene chiamato da SPFx quando il web part viene rimosso
  // dalla pagina. È importante fare unmount di React per evitare
  // memory leak — React deve poter fare pulizia dei suoi listener e stato.
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
}