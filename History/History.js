/// <reference name="MicrosoftAjax.js"/>

Type.registerNamespace("MVPSI.JAMSWeb.Controls");

MVPSI.JAMSWeb.Controls.History = function(element) {
    MVPSI.JAMSWeb.Controls.History.initializeBase(this, [element]);

}

MVPSI.JAMSWeb.Controls.History.prototype = {
    initialize: function() {
        MVPSI.JAMSWeb.Controls.History.callBaseMethod(this, 'initialize');

    },

    dispose: function() {

        MVPSI.JAMSWeb.Controls.History.callBaseMethod(this, 'dispose');
    },

    refresh: function() {
        var button = $get(this._element.id + "_DummyButton", this._element);
        button.click();
    }
};
MVPSI.JAMSWeb.Controls.History.registerClass('MVPSI.JAMSWeb.Controls.History', Sys.UI.Control);

if (typeof (Sys) !== 'undefined') Sys.Application.notifyScriptLoaded();