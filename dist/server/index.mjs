const bootstrap = ({ strapi }) => {
};
const destroy = ({ strapi }) => {
};
const register = ({ strapi }) => {
};
const config = {
  default: {},
  validator() {
  }
};
const contentTypes = {};
const ExcelJS = require("exceljs");
const controller = ({ strapi }) => ({
  async index(ctx) {
    ctx.body = strapi.plugin("export-excel").service("service").getWelcomeMessage();
  },
  async getDropDownData(ctx) {
    try {
      const excel = strapi.config.get("plugin::export-excel");
      if (!excel || typeof excel !== "object") {
        console.error("Excel config is missing or invalid.");
        throw new Error("Excel config is missing or invalid.");
      }
      let dropDownValues = [];
      const array = Object.keys(excel);
      Object.values(strapi.contentTypes).forEach((element) => {
        if (element.kind === "collectionType") {
          array.forEach((data) => {
            if (element.uid.startsWith(data)) {
              dropDownValues.push({
                label: element.info.displayName,
                value: element.uid
              });
            }
          });
        }
      });
      dropDownValues.sort((a, b) => a.label.localeCompare(b.label));
      ctx.body = {
        data: dropDownValues
      };
    } catch (error) {
      console.error("Error in getDropDownData:", error);
      ctx.throw(500, "Error fetching dropdown values");
    }
  },
  async getTableData(ctx) {
    try {
      const excel = strapi.config.get("plugin::export-excel");
      const { uid, limit, offset } = ctx.query;
      if (!excel || !excel[uid]) {
        throw new Error("Invalid or missing configuration for the requested UID.");
      }
      const query = await this.restructureObject(excel[uid], uid, limit, offset);
      const response = await strapi.entityService.findMany(uid, query);
      const header = [
        ...excel[uid].columns,
        ...Object.keys(excel[uid].relation)
      ];
      const where = excel[uid].locale === "true" ? { locale: "en" } : {};
      const count = await strapi.entityService.count(uid, where);
      const tableData = await this.restructureData(response, excel[uid]);
      ctx.body = {
        data: tableData,
        count,
        columns: header
      };
    } catch (error) {
      console.error("Error in getTableData:", error);
      ctx.throw(500, "Error fetching table data");
    }
  },
  async downloadExcel(ctx) {
    try {
      const excel = strapi.config.get("plugin::export-excel");
      const { uid } = ctx.query;
      if (!excel || !excel[uid]) {
        throw new Error("Invalid or missing configuration for the requested UID.");
      }
      const query = await this.restructureObject(excel[uid], uid);
      const response = await strapi.entityService.findMany(uid, query);
      const excelData = await this.restructureData(response, excel[uid]);
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Sheet 1");
      const headers = [
        ...excel[uid].columns,
        ...Object.keys(excel[uid].relation)
      ];
      const headerRestructure = headers.map(
        (header) => header.split("_").map((word) => word.charAt(0).toUpperCase() + word.slice(1)).join(" ")
      );
      worksheet.columns = headers.map((header, index2) => ({
        header: headerRestructure[index2],
        key: header,
        width: 20
      }));
      excelData.forEach((row) => {
        worksheet.addRow(row);
      });
      worksheet.columns.forEach((column) => {
        column.alignment = { wrapText: true };
      });
      worksheet.views = [{ state: "frozen", xSplit: 0, ySplit: 1, topLeftCell: "A" }];
      const buffer = await workbook.xlsx.writeBuffer();
      ctx.set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      ctx.body = buffer;
    } catch (error) {
      console.error("Error in downloadExcel:", error);
      ctx.throw(500, "Error generating Excel file");
    }
  },
  async restructureObject(inputObject, uid, limit, offset) {
    const excel = strapi.config.get("plugin::export-excel");
    if (!excel || !excel[uid]) {
      throw new Error("Invalid or missing configuration for the requested UID.");
    }
    const where = excel[uid].locale === "true" ? { locale: "en" } : {};
    const orderBy = { id: "asc" };
    const restructuredObject = {
      fields: inputObject.columns || "*",
      populate: {},
      where,
      orderBy,
      limit,
      offset
    };
    for (const key in inputObject.relation) {
      restructuredObject.populate[key] = {
        fields: inputObject.relation[key].column
      };
    }
    return restructuredObject;
  },
  async restructureData(data, objectStructure) {
    return data.map((item) => {
      const restructuredItem = {};
      for (const key of objectStructure.columns) {
        if (key in item) {
          restructuredItem[key] = item[key];
        }
      }
      for (const key in objectStructure.relation) {
        if (key in item) {
          const column = objectStructure.relation[key].column[0];
          if (item[key] && typeof item[key] === "object") {
            if (Array.isArray(item[key]) && item[key].length > 0) {
              restructuredItem[key] = item[key].map((obj) => obj[column]).join(" ");
            } else {
              restructuredItem[key] = item[key][column];
            }
          } else {
            restructuredItem[key] = null;
          }
        }
      }
      return restructuredItem;
    });
  }
});
const controllers = {
  controller
};
const middlewares = {};
const policies = {};
const contentAPIRoutes = [
  {
    method: "GET",
    path: "/",
    handler: "controller.index",
    config: {
      policies: []
    }
  },
  {
    method: "GET",
    path: "/dropdown-values",
    handler: "controller.getDropDownData",
    config: {
      policies: [],
      auth: false
    }
  },
  {
    method: "GET",
    path: "/table-data",
    handler: "controller.getTableData",
    config: {
      policies: [],
      auth: false
    }
  },
  {
    method: "GET",
    path: "/download-excel",
    handler: "controller.downloadExcel",
    config: {
      policies: [],
      auth: false
    }
  }
];
const routes = {
  "content-api": {
    type: "content-api",
    routes: contentAPIRoutes
  }
};
const service = ({ strapi }) => ({
  getWelcomeMessage() {
    return "Welcome to Strapi 🚀";
  }
});
const services = {
  service
};
const index = {
  bootstrap,
  destroy,
  register,
  config,
  controllers,
  contentTypes,
  middlewares,
  policies,
  routes,
  services
};
export {
  index as default
};
//# sourceMappingURL=index.mjs.map
