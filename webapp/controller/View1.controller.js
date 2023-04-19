sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/core/Core",
    "sap/m/Dialog",
    "sap/m/Button",
    "sap/m/Label",
    "sap/m/DialogType",
    "sap/m/library",
    "sap/ui/core/ValueState",
    "sap/m/MessageBox",
    "sap/m/MessageToast",
    "sap/m/TextArea",
    "sap/ui/model/FilterOperator",
    "sap/ui/model/Filter",
    "it/orogel/zcontrattoacquisto/model/Constants",
    "it/orogel/zcontrattoacquisto/model/xlsx"
],
    /**
     * @param {typeof sap.ui.core.mvc.Controller} Controller
     */
    function (Controller, Core, Dialog, Button, Label, DialogType, mobileLibrary, ValueState, MessageBox, MessageToast, TextArea, FilterOperator,
        Filter, Constants) {
        "use strict";
        // shortcut for sap.m.ButtonType
        var ButtonType = mobileLibrary.ButtonType;

        // shortcut for sap.m.DialogType
        var DialogType = mobileLibrary.DialogType;

        const oAppController = Controller.extend("it.orogel.zcontrattoacquisto.controller.View1", {
            onInit: function () {
                this.oComponent = this.getOwnerComponent();
                this.oGlobalBusyDialog = new sap.m.BusyDialog();
            },
            onScaricaTemplateDiCreazione: function (oEvent) {
                /*global XLSX*/
                const oDataExcel = [];
                for (var i = 1; i < 2; i++) {
                    var oAttributo = {
                        "Codice fornitore": "",
                        "Tipo contratto": "",
                        "Organizzazione Acquisti": "",
                        "Gruppo acquisti": "",
                        "Società": "",
                        "Data inizio val": "",
                        "Data fine val": "",
                        "Valore": "",
                        "Posizione": "",
                        "Codice articolo": "",
                        "Qtà": "",
                        "Prezzo netto": "",
                        //"Unità di prezzo": "1",
                        "Unità di prezzo": "",
                        "Divisione": "",
                        "Data inizio": "",
                        "Data fine": "",
                        "Qtà Scaglionata": "",
                        "Importo": "",
                    };
                    oDataExcel.push(oAttributo);
                }
                let wb = XLSX.utils.book_new();
                let ws = XLSX.utils.json_to_sheet(oDataExcel);
                var wscols = [{
                    wch: 15
                }, {
                    wch: 13
                }, {
                    wch: 21
                }, {
                    wch: 14
                }, {
                    wch: 7
                }, {
                    wch: 13
                }, {
                    wch: 13
                }, {
                    wch: 9
                }, {
                    wch: 9
                }, {
                    wch: 13
                }, {
                    wch: 4
                }, {
                    wch: 11
                }, {
                    wch: 13
                }, {
                    wch: 8
                }, {
                    wch: 9
                }, {
                    wch: 9
                }, {
                    wch: 14
                }, {
                    wch: 8
                }];
                ws['!cols'] = wscols;
                XLSX.utils.book_append_sheet(wb, ws, "Dati del contratto");
                XLSX.writeFile(wb, this.getView().getModel("i18n").getResourceBundle().getText("templatecreazioneexcel") + ".xlsx");
            },
            onScaricaTemplateDiModifica: function (oEvent) {
                if (!this.oSubmitDialog) {
                    this.oSubmitDialog = new Dialog({
                        type: DialogType.Message,
                        title: "Confirm",
                        content: [
                            new Label({
                                text: "Quale ordine vuoi modificare?",
                                labelFor: "submissionNote"
                            }),
                            new TextArea("submissionNote", {
                                width: "100%",
                                placeholder: "Numero Ordine",
                                liveChange: function (oEvent) {
                                    var sText = oEvent.getParameter("value");
                                    this.oSubmitDialog.getBeginButton().setEnabled(sText.length > 0);
                                }.bind(this)
                            })
                        ],
                        beginButton: new Button({
                            type: ButtonType.Emphasized,
                            text: "Submit",
                            enabled: false,
                            press: function () {
                                //var sText = Core.byId("submissionNote").getValue();
                                //MessageToast.show("Note is: " + sText);
                                this.oSubmitDialog.close();
                                this.onScaricaTemplateSelezionato(Core.byId("submissionNote").getValue());
                            }.bind(this)
                        }),
                        endButton: new Button({
                            text: "Cancel",
                            press: function () {
                                this.oSubmitDialog.close();
                            }.bind(this)
                        })
                    });
                }

                this.oSubmitDialog.open();
            },
            onScaricaTemplateSelezionato: function (NumeroContratto) {
                const oFinalFilter = new Filter({
                    filters: [],
                    and: true
                });
                var NumberFilter = [];
                NumberFilter.push(new Filter("Number", FilterOperator.EQ, NumeroContratto));
                oFinalFilter.aFilters.push(new Filter({
                    filters: NumberFilter,
                    and: false
                }));
                const oPromiseGet = new Promise((resolve, reject) => {
                    this.getView().getModel().read("/HEADERSet('" + NumeroContratto + "')", {
                        urlParameters: {
                            "$expand": "ITEMSET"
                        },
                        success: (oData) => {
                            resolve(oData);
                        },
                        error: (oData) => {
                            reject();
                        }
                    });
                });
                oPromiseGet.then(oData => {
                    /*global XLSX*/
                    const oDataExcel = [];
                    oData.ITEMSET.results.forEach(x => {
                        var oAttributo = {
                            "Numero Contratto": oData.Number,
                            "Codice fornitore": oData.Vendor,
                            "Tipo contratto": oData.Doc_Type,
                            "Organizzazione Acquisti": oData.Purch_Org,
                            "Gruppo acquisti": oData.Pur_Group,
                            "Società": oData.Comp_Code,
                            "Data inizio val": oData.Vper_Start.getDate() + "." + (oData.Vper_Start.getUTCMonth() + 1) + "." + oData.Vper_Start.getUTCFullYear(),
                            "Data fine val": oData.Vper_End.getDate() + "." + (oData.Vper_End.getUTCMonth() + 1) + "." + oData.Vper_End.getUTCFullYear(),
                            "Valore": oData.Acum_Value,
                            "Posizione": x.Item_No,
                            "Codice articolo": x.Material,
                            "Qtà": x.Target_Qty,
                            "Prezzo netto": x.Net_Price,
                            "Unità di prezzo": x.Po_Unit,
                            "Divisione": x.Plant,
                            "Data inizio": x.Valid_From.getDate() + "." + (x.Valid_From.getUTCMonth() + 1) + "." + x.Valid_From.getUTCFullYear(),
                            "Data fine": x.Valid_To.getDate() + "." + (x.Valid_To.getUTCMonth() + 1) + "." + x.Valid_To.getUTCFullYear(),
                            //"Riferimento scaglione": x.Line_No,
                            "Qtà Scaglionata": x.Scale_Base_Qty,
                            "Importo": x.Cond_Value
                        };
                        oDataExcel.push(oAttributo);
                    });
                    let wb = XLSX.utils.book_new();
                    let ws = XLSX.utils.json_to_sheet(oDataExcel);
                    var wscols = [{
                        wch: 15
                    }, {
                        wch: 15
                    }, {
                        wch: 13
                    }, {
                        wch: 21
                    }, {
                        wch: 14
                    }, {
                        wch: 7
                    }, {
                        wch: 13
                    }, {
                        wch: 13
                    }, {
                        wch: 9
                    }, {
                        wch: 9
                    }, {
                        wch: 13
                    }, {
                        wch: 4
                    }, {
                        wch: 11
                    }, {
                        wch: 13
                    }, {
                        wch: 8
                    }, {
                        wch: 9
                    }, {
                        wch: 9
                    }, {
                        wch: 9
                    }, {
                        wch: 14
                    }, {
                        wch: 8
                    }];
                    ws['!cols'] = wscols;
                    XLSX.utils.book_append_sheet(wb, ws, "Dati di aggiornamento listino");
                    XLSX.writeFile(wb, this.getView().getModel("i18n").getResourceBundle().getText("templatemodificaexcel") + ".xlsx");
                    this.oGlobalBusyDialog.close();
                }, oError => {
                    MessageBox.error(this.getView().getModel("i18n").getResourceBundle().getText("internalError"));
                    this.oGlobalBusyDialog.close();
                });
            },
            onFailedType: function () {
                sap.m.MessageToast.show(this.oComponent.i18n().getText("msg.error.filetype"));
            },
            onCaricaTemplateDiCreazione: function (oEvent) {
                this.oGlobalBusyDialog.open();
                let ofile = this.byId("fileUploader"),
                    file = ofile.oFileUpload.files[0];
                if (file && window.FileReader) {
                    let reader = new FileReader();
                    if (file.name.includes("xlsx")) {
                        reader.onload = (e) => {
                            let massiveUploadTotal = [];
                            /*global XLSX*/
                            var wb = XLSX.read(e.srcElement.result, {
                                type: 'binary'
                            }); // https://github.com/SheetJS/js-xlsx
                            wb.SheetNames.forEach((sheetName) => {
                                var roa = XLSX.utils.sheet_to_row_object_array(wb.Sheets[sheetName]);
                                var i = 1;
                                var Errore = "";
                                roa.forEach(x => {
                                    let oMassiveUpload = {};
                                    var ErroreRow = "";
                                    if (x["Codice fornitore"] !== undefined) {
                                        oMassiveUpload.CodiceFornitore = x["Codice fornitore"].toString();
                                    } else {
                                        oMassiveUpload.CodiceFornitore = "";
                                    }
                                    if (x["Tipo contratto"] !== undefined) {
                                        oMassiveUpload.TipoContratto = x["Tipo contratto"].toString();
                                    } else {
                                        oMassiveUpload.TipoContratto = "";
                                    }
                                    if (x["Organizzazione Acquisti"] !== undefined) {
                                        oMassiveUpload.OrganizzazioneAcquisti = x["Organizzazione Acquisti"].toString();
                                    } else {
                                        oMassiveUpload.OrganizzazioneAcquisti = "";
                                    }
                                    if (x["Gruppo acquisti"] !== undefined) {
                                        oMassiveUpload.GruppoAcquisti = x["Gruppo acquisti"].toString();
                                    } else {
                                        oMassiveUpload.GruppoAcquisti = "";
                                    }
                                    if (x["Società"] !== undefined) {
                                        oMassiveUpload.Societa = x["Società"].toString();
                                    } else {
                                        oMassiveUpload.Societa = "";
                                    }
                                    if (x["Data inizio val"] !== undefined) {
                                        oMassiveUpload.DataInizioVal = x["Data inizio val"].toString();
                                    } else {
                                        oMassiveUpload.DataInizioVal = "";
                                    }
                                    if (x["Data fine val"] !== undefined) {
                                        oMassiveUpload.DataFineVal = x["Data fine val"].toString();
                                    } else {
                                        oMassiveUpload.DataFineVal = "";
                                    }
                                    if (x["Valore"] !== undefined) {
                                        oMassiveUpload.Valore = x["Valore"].toString();
                                    } else {
                                        oMassiveUpload.Valore = "";
                                    }
                                    if (x["Posizione"] !== undefined) {
                                        oMassiveUpload.Posizione = x["Posizione"].toString();
                                    } else {
                                        oMassiveUpload.Posizione = "";
                                    }
                                    if (x["Codice articolo"] !== undefined) {
                                        oMassiveUpload.CodiceArticolo = x["Codice articolo"].toString();
                                    } else {
                                        oMassiveUpload.CodiceArticolo = "";
                                    }
                                    if (x["Qtà"] !== undefined) {
                                        oMassiveUpload.Qta = x["Qtà"].toString();
                                    } else {
                                        oMassiveUpload.Qta = "";
                                    }
                                    if (x["Prezzo netto"] !== undefined) {
                                        oMassiveUpload.PrezzoNetto = x["Prezzo netto"].toString();
                                    } else {
                                        oMassiveUpload.PrezzoNetto = "";
                                    }
                                    if (x["Unità di prezzo"] !== undefined) {
                                        oMassiveUpload.UnitaDiPrezzo = x["Unità di prezzo"].toString();
                                    } else {
                                        oMassiveUpload.UnitaDiPrezzo = "";
                                    }
                                    if (x["Divisione"] !== undefined) {
                                        oMassiveUpload.Divisione = x["Divisione"].toString();
                                    } else {
                                        oMassiveUpload.Divisione = "";
                                    }
                                    if (x["Data inizio"] !== undefined) {
                                        oMassiveUpload.DataInizio = x["Data inizio"].toString();
                                    } else {
                                        oMassiveUpload.DataInizio = "";
                                    }
                                    if (x["Data fine"] !== undefined) {
                                        oMassiveUpload.DataFine = x["Data fine"].toString();
                                    } else {
                                        oMassiveUpload.DataFine = "";
                                    }
                                    if (x["Qtà Scaglionata"] !== undefined) {
                                        oMassiveUpload.QtaScaglionata = x["Qtà Scaglionata"].toString();
                                    } else {
                                        oMassiveUpload.QtaScaglionata = "";
                                    }
                                    if (x["Importo"] !== undefined) {
                                        oMassiveUpload.Importo = x["Importo"].toString();
                                    } else {
                                        oMassiveUpload.Importo = "";
                                    }
                                    massiveUploadTotal.push(oMassiveUpload);
                                });
                                this._startCreazione(massiveUploadTotal);
                            });
                        };
                        reader.readAsBinaryString(file);
                    }
                }
            },
            _startCreazione: function (aFile) {
                var Header = {};
                if (aFile.length > 0) {
                    Header.Number = "X";
                    Header.Vendor = aFile[0].CodiceFornitore;
                    Header.Doc_Type = aFile[0].TipoContratto;
                    Header.Purch_Org = aFile[0].OrganizzazioneAcquisti;
                    Header.Pur_Group = aFile[0].GruppoAcquisti;
                    Header.Comp_Code = aFile[0].Societa;
                    var dateParts = aFile[0].DataInizioVal.split(".");
                    var date = new Date(+dateParts[2], dateParts[1] - 1, +dateParts[0]);
                    date.setHours(date.getHours() - date.getTimezoneOffset() / 60);
                    Header.Vper_Start = date;
                    dateParts = aFile[0].DataFineVal.split(".");
                    date = new Date(+dateParts[2], dateParts[1] - 1, +dateParts[0]);
                    date.setHours(date.getHours() - date.getTimezoneOffset() / 60);
                    Header.Vper_End = date;
                    Header.Acum_Value = aFile[0].Valore;
                }
                Header.ITEMSET = [];
                aFile.forEach(x => {
                    var Item = {};
                    Item.Number = "X";
                    Item.Item_No = x.Posizione;
                    Item.Material = x.CodiceArticolo;
                    Item.Target_Qty = x.Qta;
                    Item.Net_Price = x.PrezzoNetto;
                    Item.Po_Unit = x.UnitaDiPrezzo;
                    Item.Plant = x.Divisione;
                    if (x.DataInizio !== "") {
                        dateParts = x.DataInizio.split(".");
                        date = new Date(+dateParts[2], dateParts[1] - 1, +dateParts[0]);
                        date.setHours(date.getHours() - date.getTimezoneOffset() / 60);
                        Item.Valid_From = date;
                    }
                    if (x.DataFine !== "") {
                        dateParts = x.DataFine.split(".");
                        date = new Date(+dateParts[2], dateParts[1] - 1, +dateParts[0]);
                        date.setHours(date.getHours() - date.getTimezoneOffset() / 60);
                        Item.Valid_To = date;
                    }
                    Item.Scale_Base_Qty = x.QtaScaglionata;
                    Item.Cond_Value = x.Importo;
                    Header.ITEMSET.push(Item);
                });
                const oPromiseCreazione = new Promise((resolve, reject) => {
                    this.getView().getModel().create("/HEADERSet", Header, {
                        success: (oData) => {
                            resolve(oData);
                        },
                        error: (oData) => {
                            reject();
                        }
                    });
                });
                oPromiseCreazione.then(oData => {
                    if (typeof oData !== 'undefined') {
                        if (oData.ReturnMessage !== "" || oData.ReturnSuccessMessage !== "") {
                            if (oData.ReturnSuccessMessage !== "") {
                                MessageBox.success(this.getView().getModel("i18n").getResourceBundle().getText(oData.ReturnSuccessMessage));
                                this.oGlobalBusyDialog.close();
                            } else {
                                MessageBox.error(this.getView().getModel("i18n").getResourceBundle().getText(oData.ReturnMessage));
                                this.oGlobalBusyDialog.close();
                            }
                        }
                        else { //error during call
                            MessageBox.error(this.getView().getModel("i18n").getResourceBundle().getText("internalError"));
                            this.oGlobalBusyDialog.close();
                        }
                    } else {
                        MessageBox.error(this.getView().getModel("i18n").getResourceBundle().getText("internalError"));
                        this.oGlobalBusyDialog.close();
                    }
                }, oError => {
                    MessageBox.error(this.getView().getModel("i18n").getResourceBundle().getText("internalError"));
                    this.oGlobalBusyDialog.close();
                });
            },
            onCaricaTemplateDiModifica: function (oEvent) {
                this.oGlobalBusyDialog.open();
                let ofile = this.byId("fileUploader"),
                    file = ofile.oFileUpload.files[0];
                if (file && window.FileReader) {
                    let reader = new FileReader();
                    if (file.name.includes("xlsx")) {
                        reader.onload = (e) => {
                            let massiveUploadTotal = [];
                            /*global XLSX*/
                            var wb = XLSX.read(e.srcElement.result, {
                                type: 'binary'
                            }); // https://github.com/SheetJS/js-xlsx
                            wb.SheetNames.forEach((sheetName) => {
                                var roa = XLSX.utils.sheet_to_row_object_array(wb.Sheets[sheetName]);
                                var i = 1;
                                var Errore = "";
                                roa.forEach(x => {
                                    let oMassiveUpload = {};
                                    var ErroreRow = "";
                                    /*if (x["Numero Contratto"] !== undefined) {
                                        oMassiveUpload.NumeroContratto = x["Numero Contratto"].toString();
                                    } else {
                                        if (Errore === "") {
                                            Errore = "Numero contratto riga " + i + " non presente";
                                        } else {
                                            Errore = Errore + "Numero contratto riga " + i + " non presente";
                                        }
                                    }
                                    if (x["Posizione"] !== undefined) {
                                        oMassiveUpload.Posizione = x["Posizione"].toString();
                                    } else {
                                        if (Errore === "") {
                                            Errore = "Numero posizione riga " + i + " non presente";
                                        } else {
                                            Errore = Errore + "Numero posizione riga " + i + " non presente";
                                        }
                                    }
                                    if (x["Prezzo netto"] !== undefined) {
                                        oMassiveUpload.PrezzoNetto = x["Prezzo netto"].toString();
                                    } else {
                                        oMassiveUpload.PrezzoNetto = "";
                                    }
                                    if (x["Data inizio"] !== undefined) {
                                        oMassiveUpload.DataInizio = x["Data inizio"].toString();
                                    } else {
                                        oMassiveUpload.DataInizio = "";
                                    }
                                    if (x["Data fine"] !== undefined) {
                                        oMassiveUpload.DataFine = x["Data fine"].toString();
                                    } else {
                                        oMassiveUpload.DataFine = "";
                                    }
                                    if (x["Qtà Scaglionata"] !== undefined) {
                                        oMassiveUpload.QtaScaglionata = x["Qtà Scaglionata"].toString();
                                    } else {
                                        oMassiveUpload.QtaScaglionata = "";
                                    }
                                    if (x["Importo"] !== undefined) {
                                        oMassiveUpload.Importo = x["Importo"].toString();
                                    } else {
                                        oMassiveUpload.Importo = "";
                                    }
                                    massiveUploadTotal.push(oMassiveUpload);
                                    i++;*/
                                    if (x["Numero Contratto"] !== undefined) {
                                        oMassiveUpload.NumeroContratto = x["Numero Contratto"].toString();
                                    } else {
                                        oMassiveUpload.NumeroContratto = "";
                                    }
                                    if (x["Codice fornitore"] !== undefined) {
                                        oMassiveUpload.CodiceFornitore = x["Codice fornitore"].toString();
                                    } else {
                                        oMassiveUpload.CodiceFornitore = "";
                                    }
                                    if (x["Tipo contratto"] !== undefined) {
                                        oMassiveUpload.TipoContratto = x["Tipo contratto"].toString();
                                    } else {
                                        oMassiveUpload.TipoContratto = "";
                                    }
                                    if (x["Organizzazione Acquisti"] !== undefined) {
                                        oMassiveUpload.OrganizzazioneAcquisti = x["Organizzazione Acquisti"].toString();
                                    } else {
                                        oMassiveUpload.OrganizzazioneAcquisti = "";
                                    }
                                    if (x["Gruppo acquisti"] !== undefined) {
                                        oMassiveUpload.GruppoAcquisti = x["Gruppo acquisti"].toString();
                                    } else {
                                        oMassiveUpload.GruppoAcquisti = "";
                                    }
                                    if (x["Società"] !== undefined) {
                                        oMassiveUpload.Societa = x["Società"].toString();
                                    } else {
                                        oMassiveUpload.Societa = "";
                                    }
                                    if (x["Data inizio val"] !== undefined) {
                                        oMassiveUpload.DataInizioVal = x["Data inizio val"].toString();
                                    } else {
                                        oMassiveUpload.DataInizioVal = "";
                                    }
                                    if (x["Data fine val"] !== undefined) {
                                        oMassiveUpload.DataFineVal = x["Data fine val"].toString();
                                    } else {
                                        oMassiveUpload.DataFineVal = "";
                                    }
                                    if (x["Valore"] !== undefined) {
                                        oMassiveUpload.Valore = x["Valore"].toString();
                                    } else {
                                        oMassiveUpload.Valore = "";
                                    }
                                    if (x["Posizione"] !== undefined) {
                                        oMassiveUpload.Posizione = x["Posizione"].toString();
                                    } else {
                                        oMassiveUpload.Posizione = "";
                                    }
                                    if (x["Codice articolo"] !== undefined) {
                                        oMassiveUpload.CodiceArticolo = x["Codice articolo"].toString();
                                    } else {
                                        oMassiveUpload.CodiceArticolo = "";
                                    }
                                    if (x["Qtà"] !== undefined) {
                                        oMassiveUpload.Qta = x["Qtà"].toString();
                                    } else {
                                        oMassiveUpload.Qta = "";
                                    }
                                    if (x["Prezzo netto"] !== undefined) {
                                        oMassiveUpload.PrezzoNetto = x["Prezzo netto"].toString();
                                    } else {
                                        oMassiveUpload.PrezzoNetto = "";
                                    }
                                    if (x["Unità di prezzo"] !== undefined) {
                                        oMassiveUpload.UnitaDiPrezzo = x["Unità di prezzo"].toString();
                                    } else {
                                        oMassiveUpload.UnitaDiPrezzo = "";
                                    }
                                    if (x["Divisione"] !== undefined) {
                                        oMassiveUpload.Divisione = x["Divisione"].toString();
                                    } else {
                                        oMassiveUpload.Divisione = "";
                                    }
                                    if (x["Data inizio"] !== undefined) {
                                        oMassiveUpload.DataInizio = x["Data inizio"].toString();
                                    } else {
                                        oMassiveUpload.DataInizio = "";
                                    }
                                    if (x["Data fine"] !== undefined) {
                                        oMassiveUpload.DataFine = x["Data fine"].toString();
                                    } else {
                                        oMassiveUpload.DataFine = "";
                                    }
                                    /*if (x["Riferimento scaglione"] !== undefined) {
                                        oMassiveUpload.Line_No = x["Riferimento scaglione"].toString();
                                    } else {
                                        oMassiveUpload.Line_No = "";
                                    }*/
                                    if (x["Qtà Scaglionata"] !== undefined) {
                                        oMassiveUpload.QtaScaglionata = x["Qtà Scaglionata"].toString();
                                    } else {
                                        oMassiveUpload.QtaScaglionata = "";
                                    }
                                    if (x["Importo"] !== undefined) {
                                        oMassiveUpload.Importo = x["Importo"].toString();
                                    } else {
                                        oMassiveUpload.Importo = "";
                                    }
                                    massiveUploadTotal.push(oMassiveUpload);
                                });
                                if (Errore === "") {
                                    var NumeroContratto = [...new Set(massiveUploadTotal.map(item => item.NumeroContratto))];
                                    if (NumeroContratto.length > 1) {
                                        MessageBox.error("Modificare un solo contratto");
                                    } else {
                                        const oPromiseGet = new Promise((resolve, reject) => {
                                            this.getView().getModel().read("/HEADERSet('" + NumeroContratto[0] + "')", {
                                                urlParameters: {
                                                    "$expand": "ITEMSET"
                                                },
                                                success: (oData) => {
                                                    resolve(oData);
                                                },
                                                error: (oData) => {
                                                    reject();
                                                }
                                            });
                                        });
                                        oPromiseGet.then(oData => {
                                            var oDataExcel = [];
                                            oData.ITEMSET.results.forEach(x => {
                                                var oAttributo = {
                                                    "Numero Contratto": oData.Number,
                                                    "Codice fornitore": oData.Vendor,
                                                    "Tipo contratto": oData.Doc_Type,
                                                    "Organizzazione Acquisti": oData.Purch_Org,
                                                    "Gruppo acquisti": oData.Pur_Group,
                                                    "Società": oData.Comp_Code,
                                                    "Data inizio val": oData.Vper_Start.getDate() + "." + (oData.Vper_Start.getUTCMonth() + 1) + "." + oData.Vper_Start.getUTCFullYear(),
                                                    "Data fine val": oData.Vper_End.getDate() + "." + (oData.Vper_End.getUTCMonth() + 1) + "." + oData.Vper_End.getUTCFullYear(),
                                                    "Valore": oData.Acum_Value,
                                                    "Posizione": x.Item_No,
                                                    "Codice articolo": x.Material,
                                                    "Qtà": x.Target_Qty,
                                                    "Prezzo netto": x.Net_Price,
                                                    "Unità di prezzo": x.Po_Unit,
                                                    "Divisione": x.Plant,
                                                    "Data inizio": x.Valid_From.getDate() + "." + (x.Valid_From.getUTCMonth() + 1) + "." + x.Valid_From.getUTCFullYear(),
                                                    "Data fine": x.Valid_To.getDate() + "." + (x.Valid_To.getUTCMonth() + 1) + "." + x.Valid_To.getUTCFullYear(),
                                                    //"Riferimento scaglione": x.Line_No,
                                                    "Qtà Scaglionata": x.Scale_Base_Qty,
                                                    "Importo": x.Cond_Value
                                                };
                                                oDataExcel.push(oAttributo);
                                            });
                                            var ErroreModifica = "";
                                            //massiveUploadTotal = righe excel
                                            //oDataExcel =  righe nel documento su sap 
                                            var lengthmax = massiveUploadTotal.length;
                                            if (lengthmax > oDataExcel.length) {
                                                lengthmax = oDataExcel.length;
                                            }
                                            massiveUploadTotal.forEach(x => {
                                                if (oDataExcel[0]["Codice fornitore"] !== x.CodiceFornitore) {
                                                    ErroreModifica = "X";
                                                }
                                                if (oDataExcel[0]["Tipo contratto"] !== x.TipoContratto) {
                                                    ErroreModifica = "X";
                                                }
                                                if (oDataExcel[0]["Organizzazione Acquisti"] !== x.OrganizzazioneAcquisti) {
                                                    ErroreModifica = "X";
                                                }
                                                if (oDataExcel[0]["Gruppo acquisti"] !== x.GruppoAcquisti) {
                                                    ErroreModifica = "X";
                                                }
                                                if (oDataExcel[0]["Società"] !== x.Societa) {
                                                    ErroreModifica = "X";
                                                }
                                                if (oDataExcel[0]["Data inizio val"] !== x.DataInizioVal) {
                                                    ErroreModifica = "X";
                                                }
                                                if (oDataExcel[0]["Data fine val"] !== x.DataFineVal) {
                                                    ErroreModifica = "X";
                                                }
                                                var oFind = oDataExcel.find(y => y.Posizione === x.Posizione);
                                                if (oFind) {
                                                    if (oFind["Divisione"] !== x.Divisione) {
                                                        ErroreModifica = "X";
                                                    }
                                                }
                                            });
                                            if (ErroreModifica === "") {
                                                this._startModifica(massiveUploadTotal);
                                            } else {
                                                MessageBox.error("Sono stati modificati campi non modificabili");
                                                this.oGlobalBusyDialog.close();
                                            }
                                        }, oError => {
                                            MessageBox.error(this.getView().getModel("i18n").getResourceBundle().getText("internalError"));
                                            this.oGlobalBusyDialog.close();
                                        });

                                    }
                                } else {
                                    MessageBox.error(Errore);
                                }
                            });
                        };
                        reader.readAsBinaryString(file);
                    }
                }
            },
            _startModifica: function (aFile) {
                /*aUniqueHeader.forEach(y => {
                    var Header = {};
                    Header.Number = y;
                    Header.ITEMSET = [];
                    var items = aFile.filter(z => z.NumeroContratto === y);
                    items.forEach(y => {
                        var Item = {};
                        Item.Number = y.NumeroContratto;
                        Item.Item_No = y.Posizione;
                        Item.Net_Price = y.PrezzoNetto;
                        var dateParts = y.DataInizio.split(".");
                        var date = new Date(+dateParts[2], dateParts[1] - 1, +dateParts[0]);
                        date.setHours(date.getHours() - date.getTimezoneOffset() / 60);
                        Item.Valid_From = date;
                        dateParts = y.DataFine.split(".");
                        date = new Date(+dateParts[2], dateParts[1] - 1, +dateParts[0]);
                        date.setHours(date.getHours() - date.getTimezoneOffset() / 60);
                        Item.Valid_To = date;
                        Item.Scale_Base_Qty = y.QtaScaglionata;
                        Item.Cond_Value = y.Importo;
                        Header.ITEMSET.push(Item);
                    });
                    aHeader.push(Header);
                
                });*/
                var Header = {};
                if (aFile.length > 0) {
                    Header.Number = aFile[0].NumeroContratto;
                    Header.Vendor = aFile[0].CodiceFornitore;
                    Header.Doc_Type = aFile[0].TipoContratto;
                    Header.Purch_Org = aFile[0].OrganizzazioneAcquisti;
                    Header.Pur_Group = aFile[0].GruppoAcquisti;
                    Header.Comp_Code = aFile[0].Societa;
                    var dateParts = aFile[0].DataInizioVal.split(".");
                    var date = new Date(+dateParts[2], dateParts[1] - 1, +dateParts[0]);
                    date.setHours(date.getHours() - date.getTimezoneOffset() / 60);
                    Header.Vper_Start = date;
                    dateParts = aFile[0].DataFineVal.split(".");
                    date = new Date(+dateParts[2], dateParts[1] - 1, +dateParts[0]);
                    date.setHours(date.getHours() - date.getTimezoneOffset() / 60);
                    Header.Vper_End = date;
                    Header.Acum_Value = aFile[0].Valore;
                }
                Header.ITEMSET = [];
                aFile.forEach(x => {
                    var Item = {};
                    Item.Number = aFile[0].NumeroContratto;
                    Item.Item_No = x.Posizione;
                    Item.Material = x.CodiceArticolo;
                    Item.Target_Qty = x.Qta;
                    Item.Net_Price = x.PrezzoNetto;
                    Item.Po_Unit = x.UnitaDiPrezzo;
                    Item.Line_No = x.Line_No;
                    Item.Plant = x.Divisione;
                    if(x.DataInizio !== ""){
                        dateParts = x.DataInizio.split(".");
                        date = new Date(+dateParts[2], dateParts[1] - 1, +dateParts[0]);
                        date.setHours(date.getHours() - date.getTimezoneOffset() / 60);
                        Item.Valid_From = date;
                    }
                    if(x.DataFine !== ""){
                        dateParts = x.DataFine.split(".");
                        date = new Date(+dateParts[2], dateParts[1] - 1, +dateParts[0]);
                        date.setHours(date.getHours() - date.getTimezoneOffset() / 60);
                        Item.Valid_To = date;
                    }
                    Item.Scale_Base_Qty = x.QtaScaglionata;
                    Item.Cond_Value = x.Importo;
                    Header.ITEMSET.push(Item);
                });
                const oPromiseCreazione = new Promise((resolve, reject) => {
                    /*var batchChanges = [];
                    var sServiceUrl = "/sap/opu/odata/sap/ZCONTRATTOACQUISTO_SRV/";
                    var oDataModel = new sap.ui.model.odata.ODataModel(sServiceUrl, true);
                    for (var i = 0; i < aHeader.length; i++) {
                        batchChanges.push(oDataModel.createBatchOperation("/HEADERSet", "POST", aHeader[i]));
                    }
                    if (batchChanges.length === 0) {
                        resolve();
                    } else {
                        oDataModel.addBatchChangeOperations(batchChanges);
                        oDataModel.submitBatch(function (data, responseProcess) {
                            sap.m.MessageToast.show("Successo");
                            resolve(data);
                        }.bind(this),
                            function (err) {
                                sap.m.MessageToast.show("Errore");
                                reject();
                            });
                    }*/
                    this.getView().getModel().create("/HEADERSet", Header, {
                        success: (oData) => {
                            resolve(oData);
                        },
                        error: (oData) => {
                            reject();
                        }
                    });
                });
                oPromiseCreazione.then(oData => {
                    /*var ReturnMessage = "Errore nei seguenti contratti: ";
                    var ReturnSuccessMessage = "I seguenti contratti sono stati modificati: ";
                    var PrintMessage = "";
                    var type = false;
                    if (typeof oData !== 'undefined') {
                        oData.__batchResponses[0].__changeResponses.forEach(x => {
                            if (x.data.ReturnMessage !== "" || x.data.ReturnSuccessMessage !== "") {
                                if (x.data.ReturnSuccessMessage !== "") {
                                    if (ReturnSuccessMessage === "I seguenti contratti sono stati modificati: ") {
                                        ReturnSuccessMessage = ReturnSuccessMessage + x.data.Number;
                                    } else {
                                        ReturnSuccessMessage = ReturnSuccessMessage + ", " + x.data.Number;
                                    }
                                } else {
                                    ReturnMessage = ReturnMessage + "\n" + x.data.Number + " - " + x.ReturnMessage;
                                }
                            }
                        });
                        if (ReturnSuccessMessage !== "I seguenti contratti sono stati modificati: ") {
                            PrintMessage = PrintMessage + ReturnSuccessMessage;
                            type = true;
                        }
                        if (ReturnMessage !== "Errore nei seguenti contratti: ") {
                            PrintMessage = PrintMessage + ReturnMessage;
                        }
                        if (PrintMessage !== "") {
                            if (type === true) {
                                MessageBox.success(PrintMessage);
                                this.oGlobalBusyDialog.close();
                            } else {
                                MessageBox.error(PrintMessage);
                                this.oGlobalBusyDialog.close();
                            }
                        } else {
                            MessageBox.error(this.getView().getModel("i18n").getResourceBundle().getText("internalError"));
                            this.oGlobalBusyDialog.close();
                        }
                    }
                    else { //error during call
                        MessageBox.error(this.getView().getModel("i18n").getResourceBundle().getText("internalError"));
                        this.oGlobalBusyDialog.close();
                    }*/
                    if (typeof oData !== 'undefined') {
                        if (oData.ReturnMessage !== "" || oData.ReturnSuccessMessage !== "") {
                            if (oData.ReturnSuccessMessage !== "") {
                                MessageBox.success(this.getView().getModel("i18n").getResourceBundle().getText(oData.ReturnSuccessMessage));
                                this.oGlobalBusyDialog.close();
                            } else {
                                MessageBox.error(this.getView().getModel("i18n").getResourceBundle().getText(oData.ReturnMessage));
                                this.oGlobalBusyDialog.close();
                            }
                        }
                        else { //error during call
                            MessageBox.error(this.getView().getModel("i18n").getResourceBundle().getText("internalError"));
                            this.oGlobalBusyDialog.close();
                        }
                    } else {
                        MessageBox.error(this.getView().getModel("i18n").getResourceBundle().getText("internalError"));
                        this.oGlobalBusyDialog.close();
                    }
                }, oError => {
                    MessageBox.error(this.getView().getModel("i18n").getResourceBundle().getText("internalError"));
                    this.oGlobalBusyDialog.close();
                });
            }
        });
        return oAppController;
    });
