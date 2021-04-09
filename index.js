/* global tableau Vue */

let app = new Vue({
  el: "#app",
  data: {
    paramValues: [],
    paramValue: "",
    unregister: () => {},
    configured: false
  },
  methods: {
    // Pop open configure window
    configure: async function configure() {
      const url = `${window.location.href}/config.html`;
      await tableau.extensions.ui.displayDialogAsync(url, "", {
        width: 500,
        height: 500
      });
      this.getData();
    },
    // Pull data from worksheet and field set in settings
    getData: async function() {
      this.configured = false;
      const dashboard = tableau.extensions.dashboardContent.dashboard;
      const settings = tableau.extensions.settings.getAll();
      if (
        !settings.updateParameter ||
        !settings.listenParameter ||
        !settings.worksheet ||
        !settings.field
      )
        return console.error("Missing settings!");

      const parameter = await dashboard.findParameterAsync(
        settings.updateParameter
      );
      if (!parameter)
        return console.error(
          `Parameter ${settings.updateParameter} not found!`
        );
      let currentValue = parameter.currentValue.value;

      const worksheet = dashboard.worksheets.find(
        w => w.name === settings.worksheet
      );
      if (!worksheet)
        return console.error(`Worksheet ${settings.worksheet} not found!`);

      const data = await worksheet.getSummaryDataAsync({
        ignoreSelection: true
      });
      const field = data.columns.find(
        column => column.fieldName === settings.field
      );
      if (!field) return console.error(`Field ${settings.field} not found!`);

      let marks = data.data.map(m => m[field.index].value);
      marks = [...new Set(marks)];
      marks = marks.sort((a, b) => (a > b ? 1 : -1));
      if (marks.length === 0)
        return console.error(`No data found in ${settings.field} field!`);

      this.paramValues = marks;
      this.paramValue = marks.includes(currentValue) ? currentValue : marks[0];

      this.listen();
      this.configured = true;
    },
    // Listen for changes in listen parameter
    listen: async function() {
      this.unregister();
      const dashboard = tableau.extensions.dashboardContent.dashboard;
      const settings = tableau.extensions.settings.getAll();
      const parameter = await dashboard.findParameterAsync(
        settings.listenParameter
      );
      if (!parameter)
        return console.error(
          `Parameter ${settings.listenParameter} not found!`
        );
      this.unregister = parameter.addEventListener(
        tableau.TableauEventType.ParameterChanged,
        this.getData
      );
      this.configured = true;
    }
  },
  watch: {
    //Update parameter based on select changes
    paramValue: async function(newValue) {
      const dashboard = tableau.extensions.dashboardContent.dashboard;
      const settings = tableau.extensions.settings.getAll();
      const parameter = await dashboard.findParameterAsync(
        settings.updateParameter
      );
      if (!parameter)
        return console.error(
          `Parameter ${settings.updateParameter} not found!`
        );
      try {
        parameter.changeValueAsync(newValue);
      } catch (error) {
        console.error(`Could not update parameter ${settings.updateParameter}`);
      }
    }
  },
  created: async function() {
    // Initialize the extension and get data
    await tableau.extensions.initializeAsync({ configure: this.configure });
    this.getData();
  }
});
