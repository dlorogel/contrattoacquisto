sap.ui.define([], function () {
    "use strict";

    return {
        // All border
        oBorder: {
            bottom: {
                style: "thin"
            },
            top: {
                style: "thin"
            },
            left: {
                style: "thin"
            },
            right: {
                style: "thin"
            }
        },

        // Only top and bottom border
        oBorderBottomTop: {
            bottom: {
                style: "thin"
            },
            top: {
                style: "thin"
            }
        },
        // Stile riga Intestazione
        HEADER: {
            font: {
                bold: true,
                name: "Arial",
                sz: 10
            },
            border: {
                bottom: {
                    style: "thin"
                }
            },
            alignment: {
                horizontal: "center"
            }
        },
        HEADERROW: {
            fill: {
                fgColor: {
                    rgb: "FEC000"
                }
            },
            font: {
                bold: true,
                name: "Arial",
                sz: 10
            },
            border: {
                bottom: {
                    style: "thin"
                }
            },
            alignment: {
                horizontal: "center"
            }
        }, ROW: {
            fill: {
                fgColor: {
                    rgb: "FEC000"
                }
            },
            font: {
                bold: false,
                name: "Arial",
                sz: 10,
                color: {
                    rgb: "00CC00"
                }
            },
            border: {
                bottom: {
                    style: "thin"
                }
            },
            alignment: {
                horizontal: "center"
            }
        }
    };
});