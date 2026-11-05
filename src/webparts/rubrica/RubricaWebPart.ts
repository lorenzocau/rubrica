import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import {
  IPropertyPaneConfiguration,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle
} from "@microsoft/sp-property-pane";

import { getSP, getGraph } from "./pnpjsConfig";
import Rubrica from "./components/Rubrica";
import { IRubricaWebPartProps } from "./types";
import { risolviProprieta } from "./proprieta";
import { PAGE_SIZE_MAX, PAGE_SIZE_MIN, validaNomeLista } from "./rubricaLogic";

// BaseClientSideWebPart è la classe base di SPFx per tutti i web part.
// Il tipo generico descrive le proprietà salvate dal pannello impostazioni.
export default class RubricaWebPart extends BaseClientSideWebPart<IRubricaWebPartProps> {

  // onInit viene chiamato UNA SOLA VOLTA da SPFx prima del render.
  // È l'unico posto dove this.context è disponibile e dove dobbiamo
  // inizializzare PnPjs — da questo momento getSP() e getGraph()
  // possono essere chiamati ovunque senza passare il context.
  protected async onInit(): Promise<void> {
    await super.onInit();
    getSP(this.context);
    getGraph(this.context);
  }

  // render viene chiamato da SPFx dopo onInit (e ad ogni modifica delle impostazioni).
  // Monta il componente React principale nel DOM element che SPFx ci mette
  // a disposizione (this.domElement). risolviProprieta() completa con i default
  // le proprietà mancanti, ad esempio quelle di un web part messo in pagina
  // prima dell'introduzione delle impostazioni.
  public render(): void {
    const element = React.createElement(Rubrica, risolviProprieta(this.properties));
    ReactDom.render(element, this.domElement);
  }

  // onDispose viene chiamato da SPFx quando il web part viene rimosso
  // dalla pagina. È importante fare unmount di React per evitare
  // memory leak — React deve poter fare pulizia dei suoi listener e stato.
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  // Con true le modifiche si applicano solo premendo "Applica": evita di
  // ricaricare le liste ad ogni lettera digitata nei campi di testo.
  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: "Impostazioni della rubrica" },
          groups: [
            {
              groupName: "Generale",
              groupFields: [
                PropertyPaneTextField("titolo", { label: "Titolo (vuoto = nessun titolo)" }),
                PropertyPaneTextField("placeholderRicerca", { label: "Testo della casella di ricerca" }),
                PropertyPaneSlider("risultatiPerPagina", {
                  label: "Risultati per pagina",
                  min: PAGE_SIZE_MIN,
                  max: PAGE_SIZE_MAX,
                  step: 1,
                  showValue: true
                }),
                PropertyPaneToggle("mostraElencoCompleto", {
                  label: "Mostra tutta la rubrica senza ricerca",
                  onText: "Sì",
                  offText: "No, solo dopo una ricerca"
                })
              ]
            },
            {
              groupName: "Liste SharePoint",
              groupFields: [
                PropertyPaneTextField("listaServizi", {
                  label: "Lista servizi / interni",
                  description: "Titolo della lista o percorso (es. /sites/intranet/Lists/Rubrica). Vuoto = Rubrica",
                  onGetErrorMessage: validaNomeLista,
                  deferredValidationTime: 500
                }),
                PropertyPaneTextField("listaEsterni", {
                  label: "Lista contatti esterni",
                  description: "Titolo della lista o percorso. Vuoto = Esterni",
                  onGetErrorMessage: validaNomeLista,
                  deferredValidationTime: 500
                })
              ]
            },
            {
              groupName: "Contatti esterni",
              groupFields: [
                PropertyPaneToggle("mostraEsterni", {
                  label: "Abilita contatti esterni",
                  onText: "Attivi",
                  offText: "Disattivati"
                }),
                PropertyPaneToggle("esterniInizialmenteVisibili", {
                  label: "Checkbox esterni spuntata all'apertura",
                  onText: "Sì",
                  offText: "No"
                })
              ]
            },
            {
              groupName: "Filtri sugli utenti",
              groupFields: [
                PropertyPaneToggle("soloMembri", {
                  label: "Nascondi gli ospiti (account Guest)",
                  onText: "Nascosti",
                  offText: "Visibili"
                }),
                PropertyPaneToggle("nascondiSenzaMail", {
                  label: "Nascondi gli account senza mail",
                  onText: "Nascosti",
                  offText: "Visibili"
                })
              ]
            },
            {
              groupName: "Visualizzazione",
              groupFields: [
                PropertyPaneToggle("mostraContatore", { label: "Contatore utenti/servizi" }),
                PropertyPaneToggle("mostraPresenza", { label: "Indicatore di presenza" }),
                PropertyPaneToggle("mostraMail", { label: "Indirizzo mail" }),
                PropertyPaneToggle("mostraReparto", { label: "Reparto e ruolo" }),
                PropertyPaneToggle("mostraCellulare", { label: "Numero di cellulare" }),
                PropertyPaneToggle("linkCliccabili", { label: "Link su mail, telefono e chat Teams" })
              ]
            },
            {
              groupName: "Sviluppo e test",
              groupFields: [
                PropertyPaneToggle("usaDatiMock", {
                  label: "Usa dati di prova (non reali)",
                  onText: "Attivi",
                  offText: "Disattivati"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
