(function () {
  "use strict";

  var i18n = {
    defaultLocale: "en",
    locales: {},
    registerLocale: function (locale) {
      if (!locale || !locale.code || !locale.strings) {
        return;
      }
      this.locales[locale.code] = locale;
    }
  };

  window.PGM_I18N = i18n;
})();
