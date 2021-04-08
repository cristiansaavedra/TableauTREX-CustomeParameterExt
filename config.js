/* global tableau Vue */

let app = new Vue({
  el: "#app",
  data: {
    parameters: [],
    updateParameter: "",
    listenParameter: "",
    worksheet: "",
    worksheets: [],
    field: "",
    fields: []
  },
  methods: {
    // List all valid parameters
    getParameters: async function() {
      const dashboard = tableau.extensions.dashboardContent.dashboard;
      let parameters = await dashboard.getParametersAsync();
      this.parameters = parameters.map(p => p.name);
      const settings = tableau.extensions.settings.getAll();
      this.updateParameter =
        settings.updateParameter &&
        parameters.find(p => p.name === settings.updateParameter)
          ? settings.updateParameter
          : "";
      this.listenParameter =
        settings.listenParameter &&
        parameters.find(p => p.name === settings.listenParameter)
          ? settings.listenParameter
          : "";
    },
    // List all present worksheets
    getWorksheets: function() {
      const worksheets =
        tableau.extensions.dashboardContent.dashboard.worksheets;
      this.worksheets = worksheets.map(w => w.name);
      const settings = tableau.extensions.settings.getAll();
      let worksheetSetting =
        settings.worksheet &&
        worksheets.find(w => w.name === settings.worksheet);
      if (worksheetSetting) this.worksheet = settings.worksheet;
    },
    // Get fields present on selected worksheet
    getFields: async function(worksheetName) {
      if (!worksheetName) return (this.fields = []);
      const worksheets =
        tableau.extensions.dashboardContent.dashboard.worksheets;
      const worksheet = worksheets.find(w => w.name === worksheetName);
      const data = await worksheet.getSummaryDataAsync();
      let fields = data.columns.filter(column => column.dataType === "string");
      this.fields = fields.map(f => f.fieldName);
      const settings = tableau.extensions.settings.getAll();
      this.field =
        settings.field && fields.find(f => f.fieldName === settings.field)
          ? settings.field
          : "";
    },
    // Save the configuration to settings
    save: async function() {
      tableau.extensions.settings.set("updateParameter", this.updateParameter);
      tableau.extensions.settings.set("listenParameter", this.listenParameter);
      tableau.extensions.settings.set("worksheet", this.worksheet);
      tableau.extensions.settings.set("field", this.field);
      await tableau.extensions.settings.saveAsync();
      tableau.extensions.ui.closeDialog("");
    }
  },
  watch: {
    // Update the list of fields when the workhseet changes
    worksheet: function(worksheetName) {
      this.getFields(worksheetName);
    }
  },
  created: async function() {
    // Initialize the extension and get worksheets
    await tableau.extensions.initializeDialogAsync();
    this.getParameters();
    this.getWorksheets();
  }
});
