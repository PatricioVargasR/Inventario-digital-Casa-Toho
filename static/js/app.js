/*
  Inventario Casa Töhö — JavaScript solo para detalles visuales.
  No consulta ni guarda datos: de eso se encarga Flask con formularios normales.
*/
(function () {
  'use strict';

  // Un selector marcado con data-autoenviar envía su formulario al cambiar (filtro por proveedor)
  document.querySelectorAll('select[data-autoenviar]').forEach(selector => {
    selector.addEventListener('change', () => selector.form.submit());
  });
})();
