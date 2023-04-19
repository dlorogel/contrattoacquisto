sap.ui.define([
    "sap/ui/core/UIComponent",
    "sap/ui/Device",
    "it/orogel/zcontrattoacquisto/model/models"
],
    function (UIComponent, Device, models) {
        "use strict";

        const oAppComponent = UIComponent.extend("it.orogel.zcontrattoacquisto.Component", {
            metadata: {
                manifest: "json"
            },
            /**
             * The component is initialized by UI5 automatically during the startup of the app and calls the init method once.
             * @public
             * @override
             */
            init: function () {
                // call the base component's init function
                UIComponent.prototype.init.apply(this, arguments);

                // enable routing
                this.getRouter().initialize();

                // set the device model
                this.setModel(models.createDeviceModel(), "device");
            }
        });
        /**
 * Get resource bundle
 * -------------------
 * @returns {jQuery.sap.util.ResourceBundle|Promise|*}
 */
        oAppComponent.prototype.i18n = function () {
            return this.getModel("i18n").getResourceBundle();
        };
        oAppComponent.prototype.resetAllBusy = function () {
            Object.keys(this._oBusyControl).forEach((sId) => {
                const oControl = sap.ui.getCore().byId(sId);
                if (oControl && oControl.isBusy()) {
                    oControl.setBusy(false);
                }
            });

            sap.ui.core.BusyIndicator.hide();
        };
        return oAppComponent;
    }
);