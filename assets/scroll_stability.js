// Evita saltos de scroll durante callbacks de Dash.
(function () {
  "use strict";

  var pendingDashUpdates = 0;
  var lockedScrollY = null;
  var userScrolledDuringUpdate = false;

  window.addEventListener(
    "scroll",
    function () {
      if (pendingDashUpdates > 0) {
        userScrolledDuringUpdate = true;
      }
    },
    { passive: true }
  );

  function isDashUpdateRequest(resource) {
    try {
      var url = typeof resource === "string" ? resource : resource && resource.url;
      return typeof url === "string" && url.indexOf("_dash-update-component") !== -1;
    } catch (_e) {
      return false;
    }
  }

  function onDashUpdateStart() {
    if (pendingDashUpdates === 0) {
      lockedScrollY = window.scrollY || window.pageYOffset || 0;
      userScrolledDuringUpdate = false;
    }
    pendingDashUpdates += 1;
  }

  function onDashUpdateEnd() {
    pendingDashUpdates = Math.max(0, pendingDashUpdates - 1);
    if (pendingDashUpdates !== 0 || lockedScrollY === null) {
      return;
    }

    // Espera dos frames para que el DOM termine de pintar.
    requestAnimationFrame(function () {
      requestAnimationFrame(function () {
        if (!userScrolledDuringUpdate) {
          window.scrollTo(0, lockedScrollY);
        }
        lockedScrollY = null;
      });
    });
  }

  if (typeof window.fetch === "function") {
    var originalFetch = window.fetch.bind(window);
    window.fetch = function (resource, init) {
      var isDash = isDashUpdateRequest(resource);
      if (isDash) {
        onDashUpdateStart();
      }
      return originalFetch(resource, init).finally(function () {
        if (isDash) {
          onDashUpdateEnd();
        }
      });
    };
  }

  if (typeof window.XMLHttpRequest === "function") {
    var OriginalXHR = window.XMLHttpRequest;
    function PatchedXHR() {
      var xhr = new OriginalXHR();
      var isDash = false;

      var originalOpen = xhr.open;
      xhr.open = function (method, url) {
        isDash = typeof url === "string" && url.indexOf("_dash-update-component") !== -1;
        return originalOpen.apply(xhr, arguments);
      };

      xhr.addEventListener("loadstart", function () {
        if (isDash) {
          onDashUpdateStart();
        }
      });

      xhr.addEventListener("loadend", function () {
        if (isDash) {
          onDashUpdateEnd();
        }
      });

      return xhr;
    }

    window.XMLHttpRequest = PatchedXHR;
  }
})();
