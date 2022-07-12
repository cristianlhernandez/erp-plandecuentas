/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// eslint-disable-next-line @typescript-eslint/no-unused-vars
Office.onReady((info) => {
  // eslint-disable-next-line no-undef
  $(document).ready(async function () {
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getItemOrNullObject("Plan de Cuentas");
      await context.sync();
      // eslint-disable-next-line office-addins/load-object-before-read
      if (sheet.isNullObject) {
        let sheet = context.workbook.worksheets.add("Plan de Cuentas");
        let table = sheet.tables.add("A1:H1", true);
        table.name = "plandecuentas";
        table.getHeaderRowRange().values = [["codigo", "R", "CNC", "SR", "C", "SC", "Nombre", "Descripcion"]];
        sheet.activate();
      } else {
        CargarDatos();
      }
    });
  });
});

// Formulario Inicio
// eslint-disable-next-line no-undef
$(() => {
  // eslint-disable-next-line no-undef, @typescript-eslint/no-unused-vars
  const form = $("form")
    .dxForm({
      formData: cuenta,
      labelMode: "floating",
      minColWidth: 300,
      items: [
        {
          dataField: "codigo",
          label: { text: "Código de la Cuenta" },
          editorOptions: {
            mask: "0.0.00.00/000",
          },
          validationRules: [
            {
              type: "required",
              message: "El Código es Obligatorio",
            },
            {
              type: "async",
              message: "El Código ya existe",
              validationCallback(params) {
                return enviarRespuesta(params.value);
              },
            },
          ],
        },
        {
          dataField: "nombre",
          label: {
            text: "Nombre de la Cuenta",
          },
          validationRules: [
            {
              type: "required",
              message: "El nombre es obligatorio",
            },
          ],
        },
        {
          dataField: "descripcion",
          label: {
            text: "Descripción de la Cuenta",
          },
          editorType: "dxTextArea",
          editorOptions: {
            height: 100,
          },
        },
        {
          itemType: "button",
          horizontalAlignment: "center",
          buttonOptions: {
            text: "Nuevo Registro",
            type: "success",
            useSubmitBehavior: true,
          },
        },
      ],
    })
    .dxForm("instance");

  // eslint-disable-next-line no-undef
  $("#form-container").on("submit", async function (e) {
    e.preventDefault();
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getItem("Plan de Cuentas");
      let tabla = sheet.tables.getItem("plandecuentas");
      let codigoslice = cuenta.codigo;

      cuenta.R = codigoslice.slice(0, 1);
      cuenta.CNC = codigoslice.slice(1, 2);
      cuenta.SR = codigoslice.slice(2, 4);
      cuenta.C = codigoslice.slice(4, 6);
      cuenta.SC = codigoslice.slice(6, 9);

      let dato = [
        cuenta.codigo,
        cuenta.R,
        cuenta.CNC,
        cuenta.SR,
        cuenta.C,
        cuenta.SC,
        cuenta.nombre,
        cuenta.descripcion,
      ];

      if (cuenta.codigo !== "") {
        tabla.rows.add(null, [dato], true);
      }
      sheet.getUsedRange().format.autofitColumns();
      sheet.getUsedRange().format.autofitRows();
      await context.sync();
      CargarDatos();
    });
  });

  // Formulario Fin
});

async function CargarDatos() {
  await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Plan de Cuentas");
    let tabla = sheet.tables.getItem("plandecuentas");
    let bodyRange = tabla.getDataBodyRange().load("values");
    await context.sync();
    var bodyValues = bodyRange.values.map((value) => {
      var obj = {};
      obj.codigo = value[0];
      obj.nombre = value[6];
      obj.descripcion = value[7];

      return obj;
    });
    objetoPC = bodyValues;
    console.log(objetoPC);
  });
}

const enviarRespuesta = function (value) {
  const codigo = objetoPC.findIndex((obj) => obj.codigo == value);
  // eslint-disable-next-line no-undef
  const d = $.Deferred();
  // eslint-disable-next-line no-undef
  setTimeout(() => {
    d.resolve(codigo === -1);
  }, 1000);
  return d.promise();
};
// Data estructura.
const cuenta = {
  codigo: "",
  R: "",
  CNC: "",
  SR: "",
  C: "",
  SC: "",
  nombre: "",
  descripcion: "",
};

var objetoPC = {};
