/// <reference name="MicrosoftAjax.js"/>

Type.registerNamespace("MVPSI.JAMSWeb.Controls");

MVPSI.JAMSWeb.Controls.Monitor = function(element) {
    MVPSI.JAMSWeb.Controls.Monitor.initializeBase(this, [element]);

}

MVPSI.JAMSWeb.Controls.Monitor.prototype = {
    initialize: function() {
        MVPSI.JAMSWeb.Controls.Monitor.callBaseMethod(this, 'initialize');

    },

    dispose: function() {

        MVPSI.JAMSWeb.Controls.Monitor.callBaseMethod(this, 'dispose');
    },

    refresh: function() {
        var button = $get(this._element.id + "_DummyButton", this._element);
        button.click();
    }
};
MVPSI.JAMSWeb.Controls.Monitor.registerClass('MVPSI.JAMSWeb.Controls.Monitor', Sys.UI.Control);

if (typeof (Sys) !== 'undefined') Sys.Application.notifyScriptLoaded();